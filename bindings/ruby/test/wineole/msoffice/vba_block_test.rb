require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/vba_block'

# Stands in for COM's CodeModule. Holds text the way Excel does -- reads
# always come back with CRLF, whatever was written -- and records how many
# times the whole body was fetched, so the round-trip count can be asserted.
class FakeCodeModule
  attr_reader :reads

  # `lines:` bypasses the text split entirely, so a fake can hold blank
  # lines directly -- "\r\n".split(/\r?\n/) is [] in Ruby, not the two blank
  # lines Excel actually reports for a module emptied of its blocks.
  def initialize(text = '', lines: nil)
    @lines = lines || (text.empty? ? [] : text.split(/\r?\n/))
    @reads = 0
  end

  def CountOfLines
    @lines.length
  end

  def Lines(start, count)
    @reads += 1
    raise 'Lines(1, 0) must never be called' if count.zero?

    @lines[(start - 1), count].join("\r\n") + "\r\n"
  end

  # Excel inserts at the top, not the end.
  def AddFromString(text)
    @lines = text.split(/\r?\n/) + @lines
  end

  def DeleteLines(start, count)
    @lines.slice!(start - 1, count)
  end

  def text
    @lines.join("\n")
  end
end

class MSOfficeVBABlockTest < Minitest::Test
  B = WineOLE::MSOffice::VBABlock

  def test_a_block_is_wrapped_in_its_own_markers
    cm = FakeCodeModule.new
    B.write(cm, 'go', "Sub Go()\nEnd Sub")
    assert_includes cm.text, "'<wineole:go>"
    assert_includes cm.text, "'</wineole:go>"
    assert_includes cm.text, 'Sub Go()'
  end

  def test_rewriting_a_block_replaces_it_rather_than_stacking
    cm = FakeCodeModule.new
    B.write(cm, 'go', 'Sub Go()\nEnd Sub')
    B.write(cm, 'go', "Sub Go()\n  ' second\nEnd Sub")
    assert_equal 1, cm.text.scan("'<wineole:go>").length,
      'a second write must replace the first, not add another copy'
    assert_includes cm.text, "' second"
  end

  # VBA identifiers are case-insensitive, and so are the collection lookups
  # that hand a name back ('appform' finds AppForm). Two blocks whose names
  # differ only in case would therefore hold procedures that collide --
  # "Ambiguous name detected", and every Application.Run into the module
  # fails from then on -- so the second write must replace the first.
  def test_a_name_differing_only_in_case_replaces_the_block
    cm = FakeCodeModule.new
    B.write(cm, 'main', "Sub Go()\nEnd Sub")
    B.write(cm, 'Main', "Sub Go()\n  ' second\nEnd Sub")
    assert_equal 1, cm.text.scan(/'<wineole:main>/i).length,
      'a write whose name differs only in case must replace, not add another copy'
    assert_includes cm.text, "' second"
    refute_includes cm.text, "'<wineole:main>", 'the old block went with its markers'
  end

  def test_remove_matches_the_name_case_insensitively
    cm = FakeCodeModule.new("Sub Handwritten()\nEnd Sub")
    B.write(cm, 'main', 'Sub Go()')
    assert_equal ['Sub Handwritten()', 'End Sub'], B.remove(cm, 'MAIN')
  end

  # The whole point: the module is not ours, only the block is.
  def test_other_code_in_the_module_survives
    cm = FakeCodeModule.new("Sub Handwritten()\nEnd Sub")
    B.write(cm, 'go', "Sub Go()\nEnd Sub")
    B.write(cm, 'other', "Sub Other()\nEnd Sub")
    B.remove(cm, 'go')
    assert_includes cm.text, 'Sub Handwritten()'
    assert_includes cm.text, 'Sub Other()'
    refute_includes cm.text, 'Sub Go()'
  end

  # false when there was nothing of this name to remove; otherwise the
  # module's remaining lines, so the caller does not have to refetch the
  # body just to find out whether it is now blank.
  def test_remove_reports_whether_there_was_anything_to_remove
    cm = FakeCodeModule.new
    assert_equal false, B.remove(cm, 'go')
    B.write(cm, 'go', 'Sub Go()')
    assert_equal [], B.remove(cm, 'go')
  end

  def test_remove_hands_back_the_lines_left_over_after_the_block
    cm = FakeCodeModule.new("Sub Handwritten()\nEnd Sub")
    B.write(cm, 'go', 'Sub Go()')
    remaining = B.remove(cm, 'go')
    assert_equal ['Sub Handwritten()', 'End Sub'], remaining
  end

  # Excel reports 0 lines for a module that has never been written to, and
  # Lines(1, 0) is not a legal call.
  def test_an_empty_module_is_not_read
    cm = FakeCodeModule.new
    assert_equal false, B.remove(cm, 'go')
    assert_equal 0, cm.reads
  end

  def test_the_body_is_fetched_once_per_operation
    cm = FakeCodeModule.new("Sub A()\nEnd Sub")
    B.write(cm, 'go', 'Sub Go()')
    assert_equal 1, cm.reads, 'one Lines call, not one per line'
  end

  # Measured: a module emptied of every block still reports 2 lines of
  # "\r\n" -- CountOfLines == 2, not 0. Blank means blank after stripping,
  # not CountOfLines == 0. "\r\n".split(/\r?\n/) is [] in Ruby, so a fake
  # built from that string cannot reach this state at all; lines: holds the
  # two blank lines directly instead.
  def test_blank_sees_through_leftover_newlines
    cm = FakeCodeModule.new(lines: ['', ''])
    assert_equal 2, cm.CountOfLines
    assert_equal true, B.blank?(cm)
    assert_equal true, B.blank?(FakeCodeModule.new)
    assert_equal false, B.blank?(FakeCodeModule.new('Option Explicit'))
  end

  def test_a_name_that_would_break_the_marker_is_refused
    cm = FakeCodeModule.new
    ['a>b', "a\nb", '', 'a b'].each do |bad|
      assert_raises(ArgumentError, "#{bad.inspect} must be refused") { B.write(cm, bad, 'x') }
    end
  end

  def test_an_unclosed_marker_is_reported_rather_than_guessed
    cm = FakeCodeModule.new("'<wineole:go>\nSub Go()\nEnd Sub")
    err = assert_raises(ArgumentError) { B.remove(cm, 'go') }
    assert_match(/go/, err.message)
  end

  # A close marker with no matching open is a corrupted module, not an
  # absence -- silently reporting false here is what let the module stay
  # broken forever in the reproduction the review found.
  def test_a_close_marker_with_no_open_is_reported_as_corruption
    cm = FakeCodeModule.new("'</wineole:go>\nEnd Sub")
    err = assert_raises(ArgumentError) { B.remove(cm, 'go') }
    assert_match(/go/, err.message)
  end

  # write must refuse a payload containing a line that is itself a marker --
  # for any name, not just the one being written -- because remove cannot
  # tell it apart from a real one afterwards.
  def test_a_marker_shaped_line_in_the_payload_is_refused
    cm = FakeCodeModule.new
    err = assert_raises(ArgumentError) do
      B.write(cm, 'go', "Sub Go()\n'</wineole:go>\nEnd Sub")
    end
    assert_match(/wineole:go/, err.message)
    assert_equal '', cm.text, 'a refused write must not touch the module at all'
  end

  def test_a_marker_shaped_line_for_a_different_name_is_also_refused
    cm = FakeCodeModule.new
    assert_raises(ArgumentError) { B.write(cm, 'go', "'<wineole:other>\nEnd Sub") }
    assert_raises(ArgumentError) { B.write(cm, 'go', "'</wineole:other>\nEnd Sub") }
  end

  # Reproduces the review's exact scenario: without the fix, this used to
  # delete up to the caller's accidental marker, leave "End Sub" plus an
  # orphaned close marker behind, and make blank? never true again.
  def test_the_reproduction_from_the_review_is_now_refused_up_front
    cm = FakeCodeModule.new
    assert_raises(ArgumentError) do
      B.write(cm, 'go', "Sub Go()\n'</wineole:go>\nEnd Sub")
    end
    assert_equal true, B.blank?(cm), 'the refused write must never have touched the module'
    assert_equal false, B.remove(cm, 'go'),
      'nothing was ever added, so there is nothing to remove'
  end
end
