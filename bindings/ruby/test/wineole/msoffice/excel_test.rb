require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/excel'

# Every fake COM object here includes this, so any accidental use of
# `Application.Range` (which requires selecting a sheet first) shows up as
# a recorded call rather than as a silently-passing test. Reset in setup,
# asserted empty in teardown -- so every single test in this file doubles
# as a check that Excel#[] never calls Select.
module SelectSpy
  def self.calls
    @calls ||= []
  end

  def self.reset!
    @calls = []
  end

  def Select
    SelectSpy.calls << "#{self.class}#Select"
  end
end

# Stands in for the COM Range that a resolved lookup ends up handing to
# WineOLE::MSOffice::Range -- just enough for Range#write's shape check.
class FakeComRangeForExcel
  attr_reader :written

  def initialize(rows: 1, cols: 1)
    @rows = rows
    @cols = cols
  end

  def Rows;      Struct.new(:Count).new(@rows); end
  def Columns;   Struct.new(:Count).new(@cols); end
  def Value;     @written; end
  def Value=(v); @written = v; end
end

# Stands in for a COM Worksheet.
class FakeComWorksheetForExcel
  include SelectSpy

  attr_reader :name, :range_calls, :cells_calls, :ranges

  def initialize(name)
    @name = name
    @range_calls = []
    @cells_calls = []
    @ranges = []
  end

  def Range(addr)
    @range_calls << addr
    r = FakeComRangeForExcel.new
    @ranges << r
    r
  end

  def Cells(row, col)
    @cells_calls << [row, col]
    r = FakeComRangeForExcel.new
    @ranges << r
    r
  end
end

# Stands in for a COM Worksheets collection (either Application.Worksheets
# or Workbook.Worksheets -- both are exercised through this same fake).
class FakeComWorksheetsForExcel
  attr_reader :items, :add_after_calls

  def initialize(items)
    @items = items
    @add_after_calls = []
  end

  def Item(name_or_index)
    if name_or_index.is_a?(::Integer)
      @items.fetch(name_or_index - 1) { raise "no worksheet at index #{name_or_index}" }
    else
      @items.find { |w| w.name == name_or_index } or
        raise "no worksheet named #{name_or_index.inspect}"
    end
  end

  def Count
    @items.length
  end

  # Reads the :After keyword out of **kwargs rather than declaring it as a
  # formal `After:` parameter -- Ruby formal parameter names cannot start
  # with an uppercase letter (that parses as a constant). Excel#[] itself
  # sends this the same way any COM named argument goes out through Proxy:
  # `worksheets.Add(After: proxy)`.
  def Add(**kwargs)
    after = kwargs[:After]
    @add_after_calls << after
    new_sheet = FakeComWorksheetForExcel.new("Sheet#{@items.length + 1}")
    @items << new_sheet
    new_sheet
  end
end

# Stands in for a COM Workbook.
class FakeComWorkbookForExcel
  include SelectSpy

  attr_reader :name, :worksheets
  attr_accessor :active_sheet

  def initialize(name, worksheets:, active_sheet: nil)
    @name = name
    @worksheets = FakeComWorksheetsForExcel.new(worksheets)
    @active_sheet = active_sheet
  end

  def Worksheets
    @worksheets
  end

  def ActiveSheet
    @active_sheet
  end
end

# Stands in for a COM Workbooks collection.
class FakeComWorkbooksForExcel
  attr_reader :items, :add_calls

  def initialize(items)
    @items = items
    @add_calls = 0
  end

  def Item(name)
    @items.find { |w| w.name == name } or raise "no workbook named #{name.inspect}"
  end

  def Add
    @add_calls += 1
    wb = FakeComWorkbookForExcel.new("Book#{@items.length + 1}",
      worksheets: [FakeComWorksheetForExcel.new('Sheet1')])
    @items << wb
    wb
  end
end

# Stands in for the COM Excel.Application *and* the Proxy that wraps it --
# Excel.new takes whatever it is handed as `@ole` with no distinction, and
# every test in this file hands it one of these directly. `ole_created?`
# mirrors Proxy's own meaning: true for what .create built, false for what
# .connect attached to. `ole_release`/`ole_leave_open` stand in for the two
# Proxy bookkeeping calls Excel now delegates to (Task 9) -- deciding
# whether a release actually quits Excel is the bridge's job, done via the
# CLEANUP_STEPS a real Client#create call is handed, so this fake only
# records that the call happened rather than reimplementing that decision.
class FakeComApplication
  include SelectSpy

  attr_reader :quit_calls, :display_alerts_history, :screen_updating_history, :visible,
    :ole_release_calls, :ole_leave_open_calls
  attr_accessor :active_workbook, :active_sheet, :display_alerts, :screen_updating

  def initialize(created:, version: '11.0', workbooks: [], worksheets: [],
                 active_workbook: nil, active_sheet: nil,
                 display_alerts: true, screen_updating: true)
    @created = created
    @version = version
    @workbooks = FakeComWorkbooksForExcel.new(workbooks)
    @worksheets = FakeComWorksheetsForExcel.new(worksheets)
    @active_workbook = active_workbook
    @active_sheet = active_sheet
    @quit_calls = 0
    @display_alerts = display_alerts
    @screen_updating = screen_updating
    @display_alerts_history = []
    @screen_updating_history = []
    @ole_release_calls = 0
    @ole_leave_open_calls = 0
  end

  def ole_created?
    @created
  end

  def ole_release
    @ole_release_calls += 1
    nil
  end

  def ole_leave_open
    @ole_leave_open_calls += 1
    nil
  end

  def Version
    @version
  end

  def Workbooks
    @workbooks
  end

  def Worksheets
    @worksheets
  end

  def ActiveWorkbook
    @active_workbook
  end

  def ActiveSheet
    @active_sheet
  end

  def Quit
    @quit_calls += 1
  end

  def DisplayAlerts
    @display_alerts
  end

  def DisplayAlerts=(v)
    @display_alerts_history << v
    @display_alerts = v
  end

  def ScreenUpdating
    @screen_updating
  end

  def ScreenUpdating=(v)
    @screen_updating_history << v
    @screen_updating = v
  end

  def Visible=(v)
    @visible = v
  end
end

# Stands in for a WineOLE::Client: answers create/connect/connect_or_create
# (for Excel.create/.connect/.connect_or_create) and loopback? (for the
# Book instances Excel builds -- see Paths.convertible?). Each of the three
# also records the `cleanup:` kwarg it was handed, so a test can assert on
# exactly what Excel declared -- mirroring the real
# Client#create/connect/connect_or_create(class_name, cleanup: nil) shape.
class FakeClientForExcel
  attr_reader :create_calls, :connect_calls, :connect_or_create_calls,
    :create_cleanups, :connect_cleanups, :connect_or_create_cleanups

  def initialize(app, loopback: true)
    @app = app
    @loopback = loopback
    @create_calls = []
    @connect_calls = []
    @connect_or_create_calls = []
    @create_cleanups = []
    @connect_cleanups = []
    @connect_or_create_cleanups = []
  end

  def create(class_name, cleanup: nil)
    @create_calls << class_name
    @create_cleanups << cleanup
    @app
  end

  def connect(class_name, cleanup: nil)
    @connect_calls << class_name
    @connect_cleanups << cleanup
    @app
  end

  def connect_or_create(class_name, cleanup: nil)
    @connect_or_create_calls << class_name
    @connect_or_create_cleanups << cleanup
    @app
  end

  def loopback?
    @loopback
  end
end

class ExcelTest < Minitest::Test
  def setup
    SelectSpy.reset!
  end

  def teardown
    assert_empty SelectSpy.calls,
      'Excel must never call Select -- resolving through worksheet objects needs no active ' \
      'state, so a read must not mutate the caller\'s selection as a side effect'
  end

  def new_excel(app_obj, convert_paths: true)
    WineOLE::MSOffice::Excel.new(app_obj, client: FakeClientForExcel.new(app_obj), convert_paths: convert_paths)
  end

  # --- lifecycle: run always releases what it used ------------------------
  #
  # Whether a release also quits Excel is the bridge's decision now (it runs
  # CLEANUP_STEPS only for the last user of a record it auto-created, per
  # the cleanup: kwarg create/connect/connect_or_create declare below) --
  # not something this fake or these tests can see. What `run` itself still
  # owns, and what these pin, is that its ensure calls ole_release exactly
  # once no matter which mode built the record or what the bridge reported.

  def test_run_releases_what_it_used_regardless_of_mode
    created_app = FakeComApplication.new(created: true)
    WineOLE::MSOffice::Excel.run(:create, client: FakeClientForExcel.new(created_app)) { |xl|
      assert_instance_of WineOLE::MSOffice::Excel, xl
    }
    assert_equal 1, created_app.ole_release_calls

    attached_app = FakeComApplication.new(created: false)
    WineOLE::MSOffice::Excel.run(:connect, client: FakeClientForExcel.new(attached_app)) { |_xl| }
    assert_equal 1, attached_app.ole_release_calls
  end

  def test_run_with_connect_or_create_also_releases_regardless_of_the_bridges_report
    freshly_created = FakeComApplication.new(created: true)
    WineOLE::MSOffice::Excel.run(:connect_or_create, client: FakeClientForExcel.new(freshly_created)) { |_xl| }
    assert_equal 1, freshly_created.ole_release_calls

    already_running = FakeComApplication.new(created: false)
    WineOLE::MSOffice::Excel.run(:connect_or_create, client: FakeClientForExcel.new(already_running)) { |_xl| }
    assert_equal 1, already_running.ole_release_calls
  end

  # --- lifecycle: the block's own exception propagates -------------------

  def test_run_releases_even_when_the_block_raises
    boom = Class.new(StandardError)
    created_app = FakeComApplication.new(created: true)
    err = assert_raises(boom) do
      WineOLE::MSOffice::Excel.run(:create, client: FakeClientForExcel.new(created_app)) { raise boom, 'block failed' }
    end
    assert_equal 'block failed', err.message
    assert_equal 1, created_app.ole_release_calls
  end

  # --- ole_release / leave_open delegate to the underlying proxy ---------

  def test_ole_release_delegates_to_the_proxy
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)

    xl.ole_release
    assert_equal 1, a.ole_release_calls
  end

  def test_leave_open_delegates_to_the_proxy
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)

    xl.leave_open
    assert_equal 1, a.ole_leave_open_calls
  end

  # --- create/connect/connect_or_create declare the bridge cleanup steps --
  #
  # Quitting an auto-created Excel is the bridge's job now: it runs these
  # two steps -- suppress prompts, then Quit -- once the last user of a
  # record it auto-created releases it. All three constructors declare the
  # identical steps; see the comment on CLEANUP_STEPS and on .connect in
  # excel.rb for why that is correct even though .connect's own record is
  # (almost) never auto-created.

  def test_create_declares_displayalerts_then_quit_cleanup_steps
    app = FakeComApplication.new(created: true)
    client = FakeClientForExcel.new(app)

    WineOLE::MSOffice::Excel.create(client: client)
    assert_equal [WineOLE::MSOffice::Excel::CLEANUP_STEPS], client.create_cleanups
  end

  def test_connect_declares_the_same_cleanup_steps
    app = FakeComApplication.new(created: false)
    client = FakeClientForExcel.new(app)

    WineOLE::MSOffice::Excel.connect(client: client)
    assert_equal [WineOLE::MSOffice::Excel::CLEANUP_STEPS], client.connect_cleanups
  end

  def test_connect_or_create_declares_the_same_cleanup_steps
    app = FakeComApplication.new(created: true)
    client = FakeClientForExcel.new(app)

    WineOLE::MSOffice::Excel.connect_or_create(client: client)
    assert_equal [WineOLE::MSOffice::Excel::CLEANUP_STEPS], client.connect_or_create_cleanups
  end

  def test_cleanup_steps_are_displayalerts_off_then_quit
    assert_equal({steps: [['DisplayAlerts=', false], ['Quit']]}, WineOLE::MSOffice::Excel::CLEANUP_STEPS)
  end

  # --- create/connect/connect_or_create use the module default client ----

  def test_create_uses_the_modules_default_client_when_none_given
    require_relative '../../../lib/wineole'
    fake_app = FakeComApplication.new(created: true)
    fake_client = FakeClientForExcel.new(fake_app)
    original = WineOLE.method(:default_client)
    WineOLE.define_singleton_method(:default_client) { fake_client }

    xl = WineOLE::MSOffice::Excel.create
    assert_same fake_app, xl.ole
    assert_equal ['Excel.Application'], fake_client.create_calls
  ensure
    WineOLE.define_singleton_method(:default_client, original)
  end

  # --- resolution table: workbook part ------------------------------------

  def test_workbook_active_form_uses_active_workbook
    wb = FakeComWorkbookForExcel.new('Book1', worksheets: [FakeComWorksheetForExcel.new('Sheet1')])
    a = FakeComApplication.new(created: true, workbooks: [wb], active_workbook: wb)
    xl = new_excel(a)

    result = xl['[]']
    assert_instance_of WineOLE::MSOffice::Book, result
    assert_same wb, result.ole
  end

  def test_workbook_active_form_raises_when_nothing_is_open
    a = FakeComApplication.new(created: true, active_workbook: nil)
    xl = new_excel(a)

    err = assert_raises(RuntimeError) { xl['[]'] }
    assert_match(/active workbook/i, err.message)
  end

  def test_workbook_new_form_adds_a_workbook
    a = FakeComApplication.new(created: true, workbooks: [])
    xl = new_excel(a)

    result = xl['[:new]']
    assert_instance_of WineOLE::MSOffice::Book, result
    assert_equal 1, a.Workbooks.add_calls
    assert_same a.Workbooks.items.last, result.ole
  end

  def test_workbook_named_form_looks_up_by_name
    wb = FakeComWorkbookForExcel.new('Sales', worksheets: [FakeComWorksheetForExcel.new('Sheet1')])
    a = FakeComApplication.new(created: true, workbooks: [wb])
    xl = new_excel(a)

    result = xl['[Sales]Sheet1!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1'], wb.worksheets.items.first.range_calls
  end

  # --- resolution table: worksheet part -----------------------------------

  def test_worksheet_active_form_uses_active_sheet
    active = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, active_sheet: active)
    xl = new_excel(a)

    result = xl['!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1'], active.range_calls
  end

  def test_worksheet_active_form_raises_when_nothing_is_open
    a = FakeComApplication.new(created: true, active_sheet: nil)
    xl = new_excel(a)

    err = assert_raises(RuntimeError) { xl['!A1'] }
    assert_match(/active worksheet/i, err.message)
  end

  def test_worksheet_new_form_adds_after_the_last_sheet
    sheet1 = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, worksheets: [sheet1])
    xl = new_excel(a)

    result = xl[':new!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal [sheet1], a.Worksheets.add_after_calls
    new_sheet = a.Worksheets.items.last
    assert_equal ['A1'], new_sheet.range_calls
  end

  def test_worksheet_new_form_without_a_range_returns_the_sheet
    sheet1 = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, worksheets: [sheet1])
    xl = new_excel(a)

    result = xl[':new!']
    assert_instance_of WineOLE::MSOffice::Sheet, result
    assert_equal [sheet1], a.Worksheets.add_after_calls
  end

  def test_worksheet_first_form
    s1 = FakeComWorksheetForExcel.new('Sheet1')
    s2 = FakeComWorksheetForExcel.new('Sheet2')
    a = FakeComApplication.new(created: true, worksheets: [s1, s2])
    xl = new_excel(a)

    result = xl[':first!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1'], s1.range_calls
    assert_empty s2.range_calls
  end

  def test_worksheet_last_form
    s1 = FakeComWorksheetForExcel.new('Sheet1')
    s2 = FakeComWorksheetForExcel.new('Sheet2')
    a = FakeComApplication.new(created: true, worksheets: [s1, s2])
    xl = new_excel(a)

    result = xl[':last!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1'], s2.range_calls
    assert_empty s1.range_calls
  end

  def test_worksheet_digit_index_form
    s1 = FakeComWorksheetForExcel.new('Sheet1')
    s2 = FakeComWorksheetForExcel.new('Sheet2')
    a = FakeComApplication.new(created: true, worksheets: [s1, s2])
    xl = new_excel(a)

    result = xl['2!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1'], s2.range_calls
  end

  def test_worksheet_named_form
    s1 = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, worksheets: [s1])
    xl = new_excel(a)

    result = xl['Sheet1!A1']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1'], s1.range_calls
  end

  # --- resolution table: bare range and two-integer forms -----------------

  def test_bare_range_uses_the_active_sheet
    active = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, active_sheet: active)
    xl = new_excel(a)

    result = xl['A1:B2']
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal ['A1:B2'], active.range_calls
  end

  def test_two_integer_form_uses_the_active_sheet
    active = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, active_sheet: active)
    xl = new_excel(a)

    result = xl[2, 3]
    assert_instance_of WineOLE::MSOffice::Range, result
    assert_equal [[2, 3]], active.cells_calls
  end

  def test_two_integer_form_raises_when_nothing_is_open
    a = FakeComApplication.new(created: true, active_sheet: nil)
    xl = new_excel(a)

    err = assert_raises(RuntimeError) { xl[2, 3] }
    assert_match(/active worksheet/i, err.message)
  end

  # --- resolution table: raw-name fallback --------------------------------

  def test_raw_name_fallback_wraps_worksheets_item
    s1 = FakeComWorksheetForExcel.new('My Report')
    a = FakeComApplication.new(created: true, worksheets: [s1])
    xl = new_excel(a)

    result = xl['My Report']
    assert_instance_of WineOLE::MSOffice::Sheet, result
    assert_same s1, result.ole
  end

  # --- []= ------------------------------------------------------------

  def test_bracket_assign_writes_through_the_resolved_range
    active = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, active_sheet: active)
    xl = new_excel(a)

    xl['A1:B1'] = 7
    assert_equal 7, active.ranges.last.written
  end

  def test_bracket_assign_with_two_integers
    active = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, active_sheet: active)
    xl = new_excel(a)

    xl[2, 3] = 9
    assert_equal [[2, 3]], active.cells_calls
    assert_equal 9, active.ranges.last.written
  end

  def test_assigning_to_a_worksheet_only_address_raises
    s1 = FakeComWorksheetForExcel.new('Sheet1')
    a = FakeComApplication.new(created: true, worksheets: [s1])
    xl = new_excel(a)

    err = assert_raises(ArgumentError) { xl['Sheet1!'] = 0 }
    assert_match(/range/, err.message)
  end

  def test_assigning_to_a_workbook_only_address_raises
    wb = FakeComWorkbookForExcel.new('Book1', worksheets: [FakeComWorksheetForExcel.new('Sheet1')])
    a = FakeComApplication.new(created: true, workbooks: [wb], active_workbook: wb)
    xl = new_excel(a)

    err = assert_raises(ArgumentError) { xl['[]'] = 0 }
    assert_match(/range/, err.message)
  end

  # --- @version reaches every Sheet and Book Excel builds -----------------

  def test_version_reaches_sheet_and_book
    wb = FakeComWorkbookForExcel.new('Book1', worksheets: [FakeComWorksheetForExcel.new('Sheet1')])
    a = FakeComApplication.new(created: true, version: '12.0', workbooks: [wb],
      worksheets: [FakeComWorksheetForExcel.new('Sheet1')], active_workbook: wb)
    xl = new_excel(a)

    # XFD1 is beyond Excel 11's IV/65536 grid, but within Excel 12's -- so
    # this only succeeds if @version (not a hardcoded 11.0) reached the
    # Sheet Excel built for the lookup.
    result = xl['Sheet1!XFD1']
    assert_instance_of WineOLE::MSOffice::Range, result

    # And into the Book Excel builds, via the Sheet the Book itself builds.
    book = xl['[]']
    assert_instance_of WineOLE::MSOffice::Book, book
    s = book.sheet('Sheet1')
    refute_nil s['XFD1']
  end

  # --- no_alert / no_update restore what was there, not a hardcoded true --

  def test_no_alert_restores_a_pre_existing_false
    a = FakeComApplication.new(created: true, display_alerts: false)
    xl = new_excel(a)

    ran = false
    xl.no_alert do
      ran = true
      assert_equal false, a.DisplayAlerts
    end
    assert ran
    assert_equal false, a.DisplayAlerts,
      'no_alert must restore whatever was there before, not hardcode true the way msoffice.rb does'
  end

  def test_no_alert_restores_a_pre_existing_true
    a = FakeComApplication.new(created: true, display_alerts: true)
    xl = new_excel(a)

    xl.no_alert { assert_equal false, a.DisplayAlerts }
    assert_equal true, a.DisplayAlerts
  end

  def test_no_alert_restores_even_when_the_block_raises
    a = FakeComApplication.new(created: true, display_alerts: false)
    xl = new_excel(a)
    boom = Class.new(StandardError)

    assert_raises(boom) { xl.no_alert { raise boom } }
    assert_equal false, a.DisplayAlerts
  end

  # --- a restore that cannot happen must not become the caller's problem ---

  def gone_error
    WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800706BE)')
  end

  # An exception raised inside `ensure` REPLACES whatever the block raised.
  # Unguarded, a block that failed while the application was also going away
  # reported the cleanup instead of the real error. The fake lets the
  # suppression through and fails only on the way back, which is the order
  # the real thing fails in.
  def test_the_blocks_own_exception_survives_a_restore_that_fails
    a = FakeComApplication.new(created: false)
    def a.DisplayAlerts=(v)
      raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800706BE)') if v != false

      @display_alerts_history << v
      @display_alerts = v
    end
    xl = new_excel(a)

    boom = Class.new(StandardError)
    err = assert_raises(boom) { xl.no_alert { raise boom, 'the real problem' } }
    assert_equal 'the real problem', err.message,
      "the restore's own failure must not stand in for what the block raised"
  end

  # And with no exception of its own to defend, a failed restore must still
  # not manufacture one.
  def test_a_failing_restore_alone_does_not_raise
    a = FakeComApplication.new(created: false)
    def a.DisplayAlerts=(v)
      raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800706BE)') if v != false

      @display_alerts_history << v
      @display_alerts = v
    end
    xl = new_excel(a)

    assert_equal :fine, xl.no_alert { :fine }
  end

  # `no_alert` is still a public method any caller can wrap a Quit in by
  # hand (Quit itself no longer runs from inside this class -- the bridge
  # runs it, at release time, as one of CLEANUP_STEPS). This exercises that
  # the application dying *during* the block still leaves `no_alert`'s
  # ensure with nothing to put DisplayAlerts back on, and that this is not
  # treated as an error.
  def test_an_application_that_dies_during_the_block_is_not_an_error
    a = FakeComApplication.new(created: true)
    def a.Quit
      @quit_calls += 1
      @gone = true
    end
    def a.DisplayAlerts
      raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800706BE)') if @gone

      @display_alerts
    end
    def a.DisplayAlerts=(v)
      raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800706BE)') if @gone

      @display_alerts_history << v
      @display_alerts = v
    end
    xl = new_excel(a)

    assert_equal :done, xl.no_alert { a.Quit; :done }
    assert_equal 1, a.quit_calls, 'the Quit itself must still have happened'
  end

  # --- a failed leading read must not write nil into the flag ------------
  #
  # `previous = @ole.DisplayAlerts` is parsed before the method runs, so a
  # naive `ensure restore(:DisplayAlerts=, previous)` sees `previous == nil`
  # even when that read itself is what raised. Measured on live Excel 11:
  # writing nil to either flag sets it to *false*, silently -- so a
  # transient failure on the read would leave the flag false for the rest
  # of the session. The fix hoists the read out of the protected region so
  # a failing read raises straight to the caller with no restore attempted.

  def test_no_alert_when_the_leading_read_raises_the_caller_sees_it_and_nothing_is_written
    a = FakeComApplication.new(created: true)
    boom = Class.new(StandardError)
    def a.DisplayAlerts
      raise 'reading DisplayAlerts exploded'
    end
    xl = new_excel(a)

    err = assert_raises(RuntimeError) { xl.no_alert { flunk 'must not run the block' } }
    assert_match(/exploded/, err.message)
    assert_empty a.display_alerts_history, 'a failed read must never reach the setter'
  end

  def test_no_update_when_the_leading_read_raises_the_caller_sees_it_and_nothing_is_written
    a = FakeComApplication.new(created: true)
    def a.ScreenUpdating
      raise 'reading ScreenUpdating exploded'
    end
    xl = new_excel(a)

    err = assert_raises(RuntimeError) { xl.no_update { flunk 'must not run the block' } }
    assert_match(/exploded/, err.message)
    assert_empty a.screen_updating_history, 'a failed read must never reach the setter'
  end

  # --- no_alert / no_update require a block -------------------------------
  #
  # A bare call with no block used to perform two COM round trips and set
  # the flag false before dying as LocalJumpError from the bare `yield`.
  # Guarding at the top means it dies before touching the flag at all.

  def test_no_alert_without_a_block_raises_before_touching_the_flag
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)

    err = assert_raises(ArgumentError) { xl.no_alert }
    assert_match(/no_alert/, err.message)
    assert_match(/block/, err.message)
    assert_empty a.display_alerts_history
  end

  def test_no_update_without_a_block_raises_before_touching_the_flag
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)

    err = assert_raises(ArgumentError) { xl.no_update }
    assert_match(/no_update/, err.message)
    assert_match(/block/, err.message)
    assert_empty a.screen_updating_history
  end

  def test_no_update_restores_a_pre_existing_false
    a = FakeComApplication.new(created: true, screen_updating: false)
    xl = new_excel(a)

    xl.no_update { assert_equal false, a.ScreenUpdating }
    assert_equal false, a.ScreenUpdating
  end

  def test_no_update_restores_a_pre_existing_true
    a = FakeComApplication.new(created: true, screen_updating: true)
    xl = new_excel(a)

    xl.no_update { assert_equal false, a.ScreenUpdating }
    assert_equal true, a.ScreenUpdating
  end

  # --- show / hide ------------------------------------------------------

  def test_show_sets_visible_true
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)
    xl.show
    assert_equal true, a.visible
  end

  def test_hide_sets_visible_false
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)
    xl.hide
    assert_equal false, a.visible
  end

  # --- passthrough ------------------------------------------------------

  def test_ole_reader_exposes_the_underlying_proxy
    a = FakeComApplication.new(created: true)
    xl = new_excel(a)
    assert_same a, xl.ole
  end

  def test_unknown_methods_go_to_com
    a = FakeComApplication.new(created: true, version: '11.0')
    xl = new_excel(a)
    assert_equal '11.0', xl.Version
  end

  # --- the wrapper claims nothing in the root namespace --------------------

  # The inverse of what an earlier draft asserted: requiring the whole
  # wrapper must leave the root namespace untouched. Reaching into it from
  # a library is intrusive, and the alias that used to live there brought
  # two failure modes with it -- a silent no-op when something else had
  # already defined MSOffice, and a result that depended on process-wide
  # load order.
  def test_requiring_the_wrapper_defines_no_root_level_constant
    require_relative '../../../lib/wineole/msoffice'

    refute defined?(MSOffice), 'the wrapper must not claim a root-level MSOffice'
    assert_equal WineOLE::MSOffice::Excel, WineOLE::MSOffice.const_get(:Excel)
  end
end
