require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/vba'

class MSOfficeVBATest < Minitest::Test
  V = WineOLE::MSOffice::VBA

  # codepage memoizes, so without this a test that simulates a broken
  # registry would read a value an earlier test had already cached and
  # pass or fail on test order rather than on what it is testing.
  def setup
    V.forget_codepage
  end

  def teardown
    V.forget_codepage
  end

  # Replaces the one method that shells out, so no test touches the real
  # registry. Returns [stdout, success].
  #
  # Not implemented with minitest/mock's Object#stub: on this host that
  # module lives only inside the minitest 5.x gem, and `minitest/autorun`
  # (from minitest 6.0.6, installed alongside it) has already activated
  # minitest 6 by the time a test would require it, so `require
  # 'minitest/mock'` fails outright (the separately-packaged
  # `minitest-mock` gem ships no code on this host -- it is an empty
  # placeholder). A plain singleton-method swap needs no extra library and
  # gives the exact same behavior.
  def stub_reg(result)
    calls = []
    original = V.method(:run_reg)
    V.define_singleton_method(:run_reg) { |*args| calls << args; result }
    yield
    calls
  ensure
    V.define_singleton_method(:run_reg, original)
  end

  def test_state_reads_the_dword
    stub_reg(["\nHKEY_CURRENT_USER\\...\n    AccessVBOM    REG_DWORD    0x1\n", true]) do
      assert_equal :enabled, V.state
      assert_equal true, V.enabled?
    end
  end

  def test_state_when_disabled
    stub_reg(["    AccessVBOM    REG_DWORD    0x0\n", true]) do
      assert_equal :disabled, V.state
      refute V.enabled?
    end
  end

  # A missing value and a missing key both exit non-zero, and wine writes the
  # explanation to stdout in its own language -- so only the exit status can
  # be trusted.
  def test_state_when_the_value_is_absent
    stub_reg(['reg: <a localized not-found message>', false]) do
      assert_equal :unset, V.state
      refute V.enabled?
    end
  end

  # Measured: wine indents the line and leaves a CR on the end, so the raw
  # field is neither "932" nor at the index you would expect. Stripping the
  # whole line before splitting is what fixes both -- drop that and this
  # test fails, which is how it was confirmed to be the load-bearing part.
  # (A second test once duplicated this one under a different name, asserting
  # only that `Encoding.find` did not raise -- which it never can on its own
  # terms. That coverage is folded in here rather than dropped: this is the
  # one test carrying the "indented line, trailing CR, still a usable name"
  # property, via the exact-string assertion below.)
  def test_the_codepage_is_read_from_the_registry
    stub_reg(["    ACP    REG_SZ    932\r\n", true]) do
      assert_equal 'CP932', V.codepage
    end
  end

  # `read` picks its line by matching the NAME field exactly, not by
  # substring -- so a value whose *data* happens to contain another value's
  # name must not be mistaken for that value's line.
  def test_state_is_not_fooled_by_the_name_appearing_inside_other_data
    # "NameAccessVBOM" contains the value name as a substring, but its
    # leading "N" is not a hex digit, so mistaking this line for the real
    # one would parse as 0 -- :disabled -- while the real line says enabled.
    stub_reg(["    SomethingElse    REG_SZ    NameAccessVBOM\r\n    AccessVBOM    REG_DWORD    0x1\r\n", true]) do
      assert_equal :enabled, V.state
    end
  end

  # More than one line naming the value means the single-line shape this
  # parser assumes has stopped holding -- that is a bug to surface, not a
  # state to quietly resolve by taking the first match.
  def test_state_raises_when_more_than_one_line_names_the_value
    stub_reg(["    AccessVBOM    REG_DWORD    0x1\r\n    AccessVBOM    REG_DWORD    0x0\r\n", true]) do
      assert_raises(WineOLE::MSOffice::VBA::Error) { V.state }
    end
  end

  # The command can succeed (exit 0) yet produce output this parser cannot
  # make sense of -- that is a parser bug or a wine output change, and must
  # not be reported as :unset (which means "the user never configured it").
  def test_state_raises_rather_than_reporting_unset_when_output_does_not_parse
    stub_reg(["some unrelated line that never names the value\r\n", true]) do
      assert_raises(WineOLE::MSOffice::VBA::Error) { V.state }
    end
  end

  def test_state_is_unset_only_when_the_command_says_the_value_is_not_there
    stub_reg(['reg: <a localized not-found message>', false]) do
      assert_equal :unset, V.state
    end
  end

  def test_an_unreadable_codepage_raises_rather_than_guessing
    stub_reg(['reg: <not found>', false]) do
      err = assert_raises(WineOLE::MSOffice::VBA::Error) { V.codepage }
      assert_match(/ACP/, err.message)
    end
  end

  def test_an_unknown_codepage_raises
    stub_reg(["    ACP    REG_SZ    99999\r\n", true]) do
      assert_raises(WineOLE::MSOffice::VBA::Error) { V.codepage }
    end
  end

  def test_enable_and_disable_pass_the_right_arguments
    calls = stub_reg(['', true]) { assert_equal true, V.enable! }
    assert_equal 1, calls.length
    args = calls.first.flatten
    assert_equal 'add', args[0]
    assert_includes args, 'AccessVBOM'
    assert_includes args, 'REG_DWORD'
    assert_equal '1', args[args.index('/d') + 1]

    calls = stub_reg(['', true]) { V.disable! }
    args = calls.first.flatten
    assert_equal '0', args[args.index('/d') + 1]
  end

  # wine reg add prints a localized success message to stdout. Only the exit
  # status says whether it worked.
  def test_enable_reports_failure_by_exit_status_not_by_output
    stub_reg(['<some localized text>', true]) { assert_equal true, V.enable! }
    stub_reg(['<some localized text>', false]) { assert_equal false, V.enable! }
  end
  # codepage costs a `wine reg` subprocess -- 328 ms measured on this host --
  # and an import used to pay it twice. The machine's ANSI codepage cannot
  # change while a process runs, so reading it once is not a staleness risk.
  def test_the_codepage_is_read_from_the_registry_only_once
    V.forget_codepage
    calls = 0
    original = V.method(:read)
    V.define_singleton_method(:read) { |*a| calls += 1; original.call(*a) }
    begin
      first = V.codepage
      assert_equal first, V.codepage
      assert_equal first, V.codepage
      assert_equal 1, calls, 'three calls, one registry read'
    ensure
      V.define_singleton_method(:read, original)
      V.forget_codepage
    end
  end

end
