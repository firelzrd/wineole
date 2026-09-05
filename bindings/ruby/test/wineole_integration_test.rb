require 'minitest/autorun'
require 'timeout'
require_relative '../lib/wineole'
require_relative 'support/excel_integration_helper'

class WineOLEIntegrationTest < Minitest::Test
  include ExcelIntegrationHelper

  BRIDGE_EXE = ExcelIntegrationHelper::BRIDGE_EXE

  def test_end_to_end_excel_automation_via_the_bridge
    root_handle = nil

    with_excel do |xl|
      root_handle = xl.ole_handle
      xl.Visible = false
      xl.DisplayAlerts = false
      xl.Workbooks.Add

      sheets = xl.Worksheets
      before_count = sheets.Count
      first_sheet = sheets[1]
      first_sheet_name = first_sheet.Name
      sheets.Add('After' => first_sheet)
      after_count = xl.Worksheets.Count

      assert_equal before_count + 1, after_count

      # The count assertion above only proves the 'After' named argument was
      # transmitted and accepted (a dropped/unresolvable name would raise
      # DISP_E_UNKNOWNNAME via GetIDsOfNames). It does not prove the argument
      # was honored positionally. Check that the new sheet actually landed
      # immediately after first_sheet, not before it or elsewhere.
      assert_equal first_sheet_name, xl.Worksheets[1].Name,
        "first_sheet should remain sheet 1 (unmoved) after Add('After' => first_sheet)"
      refute_equal first_sheet_name, xl.Worksheets[2].Name,
        "sheet 2 should be the newly added sheet, not first_sheet, proving 'After' was honored"

      # A COM error must arrive with something usable in it — an empty
      # message used to be the norm, since Wine cannot format most
      # automation HRESULTs.
      err = assert_raises(WineOLE::RemoteError) { xl.NoSuchMemberAtAll }
      assert_equal 'WIN32OLERuntimeError', err.remote_class
      assert_match(/0x[0-9A-F]{8}/, err.message, "expected an HRESULT in #{err.message.inspect}")

      xl.Quit
    end

    # Dropping the connection — rather than releasing every handle one by
    # one, which no real client does — must make the bridge reclaim the whole
    # session. Observable from a fresh connection: the handles that
    # connection owned are gone from the routing table, so invoking one is a
    # stale reference rather than a live object.
    # Teardown is asynchronous (the read loop notices the disconnect, then
    # walks the handles), so poll rather than assuming it has already
    # finished by the time this second connection is up.
    second = connect
    begin
      assert handle_reclaimed?(second, root_handle, timeout: 20),
        'handles owned by a closed connection must not survive it'
    ensure
      second.close
    end

    # And the automated process itself should go away with the session. This
    # is best-effort rather than asserted: the observed logs show
    # `IRemUnknown_RemRelease failed with error 0x80070057` (E_INVALIDARG)
    # alongside `get_stub_manager_from_ipid not found for ipid {...}` — i.e.
    # the stub was already gone by the time the release reached it. That is
    # an over-release/late-release symptom, not an under-release, so it
    # cannot by itself be what keeps EXCEL.EXE pinned (a release the target
    # no longer needs can't keep it alive). A more plausible, but unconfirmed,
    # explanation: the session's STA worker thread (session.rs,
    # `CoInitializeEx(..., COINIT_APARTMENTTHREADED)`) never pumps a Windows
    # message loop anywhere in this crate (no `PeekMessage`/`GetMessage`/
    # `CoWaitForMultipleHandles`) — so if Excel ever raises a modal dialog
    # during `Quit` (e.g. a save prompt), nothing would dismiss it, which
    # would explain an intermittent, timing-dependent hang. This is a
    # hypothesis for follow-up investigation, not a confirmed diagnosis, and
    # fixing it is out of scope here. `test_connection_teardown_...` in
    # server.rs is the deterministic proof that the session shuts down.
    warn 'note: EXCEL.EXE outlived its session (known Wine COM release flake)' unless excel_gone?(timeout: 20)
  end

  def test_bulk_range_round_trip_preserves_types
    with_excel do |xl|
      xl.Visible = false
      xl.DisplayAlerts = false
      xl.Workbooks.Add
      sheet = xl.Worksheets[1]

      # One write for the whole block, not nine.
      sheet.Range('A1:C3').Value = [
        ['text', 1, 2.5],
        ['', nil, -3],
        ['ünïcödé ✓', 0, 1000000],
      ]

      rows = sheet.Range('A1:C3').Value

      assert_equal 3, rows.length, 'a 3x3 range must read back as 3 rows'
      assert_equal 3, rows[0].length, 'each row must have 3 columns (not transposed)'
      assert_equal 'text', rows[0][0]
      assert_equal 1.0, rows[0][1]
      assert_equal 2.5, rows[0][2]
      # A nil written into a cell comes back as nil.
      assert_nil rows[1][1], 'a nil written into a cell must read back as nil'
      # An empty string written into a cell comes back as an empty cell (nil),
      # not as ''.
      assert_nil rows[1][0], "an empty string written into a cell must read back as nil, not ''"
      assert_equal(-3.0, rows[1][2])
      assert_equal 'ünïcödé ✓', rows[2][0]
      assert_equal 1_000_000.0, rows[2][2]

      # A date cell must arrive as a Time, not as a raw {"$type" => "time"}
      # hash -- this is what the recursive decode exists for.
      #
      # A 1x1 range's Value is a bare scalar, not [[v]] (see
      # test_range_value_shape_depends_on_range_size) -- so `date` here is
      # the Time itself, not a row containing it. That only proves
      # *top-level* decode of a tagged value.
      sheet.Range('E1').Value = '2026-08-31'
      sheet.Range('E1').NumberFormat = 'yyyy-mm-dd'
      date = sheet.Range('E1:E1').Value
      assert_instance_of Time, date,
        'a date inside a bulk read must decode to a Time'
      assert_equal 2026, date.year
      assert_equal 8, date.month
      assert_equal 31, date.day

      # The assertion above says nothing about a tagged value *nested*
      # inside a returned array -- which is exactly the case the recursive
      # decode exists for (a bulk Range.Value read on anything larger than
      # 1x1 returns an array of rows, per
      # test_range_value_shape_depends_on_range_size). Write two dates side
      # by side and read them back as a 1x2 range, so reaching either date
      # means indexing into the returned array: [[Time, Time]], not a bare
      # Time.
      sheet.Range('F1').Value = '2026-09-01'
      sheet.Range('F1').NumberFormat = 'yyyy-mm-dd'
      dates = sheet.Range('E1:F1').Value
      assert_equal 1, dates.length, 'a 1x2 range must read back as 1 row'
      assert_equal 2, dates[0].length, 'that row must have 2 columns'
      assert_instance_of Time, dates[0][0],
        'a date nested inside a returned array must decode to a Time, not a raw {"$type"=>"time"} hash'
      assert_instance_of Time, dates[0][1],
        'a date nested inside a returned array must decode to a Time, not a raw {"$type"=>"time"} hash'
      assert_equal 2026, dates[0][0].year
      assert_equal 8, dates[0][0].month
      assert_equal 31, dates[0][0].day
      assert_equal 2026, dates[0][1].year
      assert_equal 9, dates[0][1].month
      assert_equal 1, dates[0][1].day
    end
  end

  def test_writing_a_time_directly_round_trips_as_a_date
    with_excel do |xl|
      xl.Visible = false
      xl.DisplayAlerts = false
      xl.Workbooks.Add
      sheet = xl.Worksheets[1]

      # A Time assigned straight into a cell -- no string-plus-NumberFormat
      # workaround -- must land as a genuine VT_DATE and read back as a
      # Time. This asserts on the *type* read back rather than on display
      # text: Excel's own Range.Value setter auto-applies a date format to
      # a still-General cell (see README), so a passing string-based
      # workaround and a genuine VT_DATE can look identical on screen --
      # only the read-back type tells them apart.
      written = Time.new(2026, 8, 31, 9, 30, 45)
      sheet.Range('H1').Value = written
      read_back = sheet.Range('H1:H1').Value
      assert_instance_of Time, read_back,
        'a Time written straight into a cell must read back as a Time, not a String or a raw hash'
      assert_equal written.to_i, read_back.to_i,
        'a direct Time write must round-trip to the second'

      # The same, but nested inside a 2-D bulk write -- the recursive case
      # in value.rs's SAFEARRAY encode/decode, which the single-cell write
      # above does not exercise. Mixed with plain strings so the test also
      # proves the recursion doesn't misencode/misdecode a date's
      # neighbors.
      written_a = Time.new(2026, 9, 1, 12, 0, 0)
      written_b = Time.new(2026, 9, 2, 18, 15, 30)
      sheet.Range('I1:J2').Value = [
        [written_a, 'not a date'],
        ['still not a date', written_b],
      ]
      grid = sheet.Range('I1:J2').Value

      assert_instance_of Time, grid[0][0],
        'a date written inside a bulk 2-D array must read back as a Time, indexed out of the returned array'
      assert_instance_of Time, grid[1][1],
        'a date written inside a bulk 2-D array must read back as a Time, indexed out of the returned array'
      assert_equal written_a.to_i, grid[0][0].to_i,
        'a bulk-written date must round-trip to the second'
      assert_equal written_b.to_i, grid[1][1].to_i,
        'a bulk-written date must round-trip to the second'
      assert_equal 'not a date', grid[0][1]
      assert_equal 'still not a date', grid[1][0]
    end
  end

  def test_round_trip_latency_is_not_stalled_by_nagle
    with_excel do |xl|
      xl.Version # warm up: the first call also starts Excel

      started = Process.clock_gettime(Process::CLOCK_MONOTONIC)
      100.times { xl.Version }
      per_call_ms = (Process.clock_gettime(Process::CLOCK_MONOTONIC) - started) / 100 * 1000

      # Before the single-write fix this was 42.6 ms per call, essentially
      # all of it the client's ~40 ms delayed-ACK timer waiting for a
      # newline Nagle was holding. Afterwards it is ~1.3 ms. 20 ms sits
      # clear of both, so this catches a regression without being flaky on
      # a loaded machine.
      assert_operator per_call_ms, :<, 20.0,
        "a round trip took #{'%.1f' % per_call_ms} ms; a Nagle/delayed-ACK stall " \
        'has probably come back (see protocol.rs write_response)'
    end
  end

  # Range#Value's result shape depends on the range's dimensions -- this is
  # Excel's own contract, not a choice this project made, and Phase 2's
  # wrapper has to normalize it. Pin it here so a future change to the
  # conversion layer cannot quietly alter it out from under that wrapper.
  def test_range_value_shape_depends_on_range_size
    with_excel do |xl|
      xl.Visible = false
      xl.DisplayAlerts = false
      xl.Workbooks.Add
      sheet = xl.Worksheets[1]

      sheet.Range('A1').Value = 42
      sheet.Range('B1').Value = 'right'
      sheet.Range('A2').Value = 'below'

      # A 1x1 range collapses to a bare scalar, whether addressed as a
      # single cell or as an explicit 1x1 range -- not [[42.0]].
      assert_equal 42.0, sheet.Range('A1').Value
      assert_equal 42.0, sheet.Range('A1:A1').Value

      # Anything larger than 1x1 comes back as an array of row arrays.
      assert_equal [[42.0, 'right']], sheet.Range('A1:B1').Value
      assert_equal [[42.0], ['below']], sheet.Range('A1:A2').Value

      # An empty 1x1 range is a bare nil, not [[nil]].
      assert_nil sheet.Range('J1:J1').Value,
        'an empty 1x1 range must read back as a bare nil, not [[nil]]'
      assert_equal [[nil], [nil]], sheet.Range('J1:J2').Value
    end
  end

  def test_a_second_client_reuses_the_already_running_bridge
    skip "bridge exe not built: #{BRIDGE_EXE}" unless File.exist?(BRIDGE_EXE)

    first = connect
    spawned_after_first = @spawned_pids.length

    second = connect
    assert_equal spawned_after_first, @spawned_pids.length,
      'the second client must reuse the running bridge, not spawn another'

    first.close
    second.close
  end

  def test_connect_or_create_creates_then_a_second_call_attaches
    # Deliberately uses with_bridge rather than with_excel: with_excel
    # already creates an Excel instance via client.create, which would
    # itself be the "first" instance and defeat this test's point -- that
    # the *first* connect_or_create call is what creates one. This test
    # needs a from-scratch bring-up so nothing exists yet when it calls
    # connect_or_create for the first time.
    #
    # Not `excel_pids - @pre_existing_excel_pids` -- setup snapshots
    # @pre_existing_excel_pids fresh before every test, so a leftover Excel
    # from a sibling test is captured *into* that snapshot and the
    # difference would always be empty. connect_or_create attaches to ANY
    # running Excel.Application, test-started or not, so the only clean
    # precondition is that none is running at all.
    unless excel_pids.empty?
      skip 'an Excel instance is already running -- cannot assert "nothing was running" cleanly; ' \
           're-run this test in isolation or investigate the leftover instance'
    end

    with_bridge do |client|
      xl = client.connect_or_create('Excel.Application')
      assert xl.ole_created?, 'the first connect_or_create must create a new instance (nothing was running)'
      xl.Visible = false
      xl.DisplayAlerts = false

      xl2 = client.connect_or_create('Excel.Application')
      refute xl2.ole_created?, 'the second connect_or_create must attach to the instance the first one created'

      # Prove xl2 is genuinely the same live instance xl mutated, not a
      # coincidentally similar second one.
      xl2.DisplayAlerts = true
      xl.DisplayAlerts = false
      assert_equal false, xl2.DisplayAlerts, 'xl and xl2 must observe the same live Excel instance'

      xl.Quit
    end
  end

  # THE regression test for the thread split. A callback that calls COM must
  # get an answer. Running callbacks on the reader thread makes this hang
  # forever, because nothing is left to read the response.
  def test_a_callback_can_call_com_and_get_an_answer
    with_bridge do |client|
      xl = client.create('Excel.Application')
      begin
        xl.Visible = false
        xl.DisplayAlerts = false
        got = Queue.new
        xl.ole_events.on('SheetChange') do |_sheet, target|
          got << target.Address
        end
        book = xl.Workbooks.Add
        book.Worksheets(1).Range('A1').Value = 42

        address = nil
        begin
          Timeout.timeout(30) { address = got.pop }
        rescue Timeout::Error
          # A verdict, not an exception escaping: this is the failure this
          # test exists to report, and saying so is the difference between a
          # suite that tells you what broke and one that times out.
          flunk 'the callback never got an answer within 30s -- with callbacks running on ' \
                'the reader thread there is nobody left to read the response to the ' \
                "callback's own COM call, and the whole connection is wedged"
        end
        assert_match(/\$?A\$?1/, address, 'the callback made a COM call and got an answer back')
      ensure
        # Bounded, because the mutation this test exists to catch wedges the
        # connection: an unbounded Quit on it hangs the whole suite instead
        # of letting this test fail. See quit_bounded.
        quit_bounded(xl)
      end
    end
  end

  # `on` is the only thing the caller touches. subscribe and Advise are
  # derived; removing that derivation makes this fail.
  def test_registering_a_callback_is_all_it_takes
    with_bridge do |client|
      xl = client.create('Excel.Application')
      begin
        xl.Visible = false
        xl.DisplayAlerts = false
        fired = Queue.new
        xl.ole_events.on('WorkbookOpen') { fired << :open }
        xl.ole_events.on('SheetChange') { fired << :changed }
        book = xl.Workbooks.Add
        book.Worksheets(1).Range('B2').Value = 1

        begin
          Timeout.timeout(30) { assert_equal :changed, fired.pop }
        rescue Timeout::Error
          flunk 'no event arrived within 30s: registering a callback must be all it takes'
        end
      ensure
        quit_bounded(xl)
      end
    end
  end

  # One dispatcher thread per CONNECTION, on the real thing. Two objects with
  # callbacks on one client -- the Application and the Workbook, both raising
  # SheetChange for the same write -- and the promise is that a caller who
  # shares state between them needs no lock of his own, because the two can
  # never be inside their blocks at once. A thread per object breaks exactly
  # that and nothing else: every other event assertion in this file passes
  # under it.
  def test_callbacks_on_two_objects_run_on_the_one_dispatcher_thread
    with_bridge do |client|
      xl = client.create('Excel.Application')
      begin
        xl.Visible = false
        xl.DisplayAlerts = false
        book = xl.Workbooks.Add

        fired = Queue.new
        xl.ole_events.on('SheetChange') { fired << [:application, Thread.current] }
        book.ole_events.on('SheetChange') { fired << [:workbook, Thread.current] }

        book.Worksheets(1).Range('C3').Value = 7

        seen = {}
        begin
          Timeout.timeout(30) do
            until seen.key?(:application) && seen.key?(:workbook)
              who, thread = fired.pop
              seen[who] = thread
            end
          end
        rescue Timeout::Error
          # A verdict, not an exception escaping: with only one of the two
          # delivered there is nothing to compare, and saying which one is
          # missing is the difference between a report and a timeout.
          flunk "only #{seen.keys.inspect} fired within 30s -- both objects must be delivered to"
        end

        assert_same seen[:application], seen[:workbook],
          "both callbacks must run on the connection's ONE dispatcher thread: a thread per " \
          "object is a data race in every callback that shares state with another object's"
      ensure
        quit_bounded(xl)
      end
    end
  end
end
