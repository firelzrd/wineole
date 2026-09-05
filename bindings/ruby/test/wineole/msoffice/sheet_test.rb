require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/sheet'
require_relative '../../../lib/wineole/msoffice/vba'

# Stands in for a VBA CodeModule, just enough for VBABlock to read and
# write it. Kept in sync by hand with the fake of the same shape in
# book_test.rb -- see that file's FakeCodeModule for why `lines:` bypasses
# the text split.
class FakeCodeModuleForSheet
  def initialize(text = '', lines: nil)
    @lines = lines || (text.empty? ? [] : text.split(/\r?\n/))
  end

  def CountOfLines = @lines.length

  def Lines(start, count)
    raise 'Lines(1, 0) must never be called' if count.zero?

    @lines[(start - 1), count].join("\r\n") + "\r\n"
  end

  def AddFromString(t) = @lines = t.split(/\r?\n/) + @lines
  def DeleteLines(start, count) = @lines.slice!(start - 1, count)
  def text = @lines.join("\n")
end

# Stands in for a VBComponent inside a workbook's VBProject -- a
# worksheet's own module, reached through Sheet's Parent/VBProject chain.
class FakeVBComponentForSheet
  attr_accessor :Name

  def initialize(name)
    @Name = name
    @code_module = FakeCodeModuleForSheet.new
  end

  def CodeModule = @code_module
end

class FakeVBComponentsForSheet
  def initialize(project) = @project = project

  def Item(name)
    @project.component_list.find { |c| c.Name == name } ||
      raise(WineOLE::RemoteError.new('X', 'not found'))
  end
end

# Stands in for a workbook's VBProject, reached from a sheet through
# `Parent.VBProject` -- CodeName names a module that already exists (a
# worksheet's own module cannot be created or deleted, only found).
class FakeVBProjectForSheet
  attr_reader :component_list

  def initialize
    @component_list = []
  end

  # A Hash view keyed by current Name -- mirrors FakeVBProject#components
  # in book_test.rb.
  def components
    @component_list.each_with_object({}) { |c, h| h[c.Name] = c }
  end

  def VBComponents = @vb_components ||= FakeVBComponentsForSheet.new(self)

  def add_existing(name)
    c = FakeVBComponentForSheet.new(name)
    @component_list << c
    c
  end
end

# Stands in for the Workbook a worksheet's Parent answers -- just enough
# for Sheet#own_code_module to reach VBProject through it. `denied:` makes
# VBProject raise the way a real Excel does when the registry refuses
# access -- own_code_module's rescue is otherwise untested, since none of
# the other fakes in this file ever raise on Parent/VBProject.
# A project that opens fine but whose component lookup fails. Needed to
# tell "AccessVBOM is off" apart from "that component is not there": the
# rescue used to cover both and answered every one of them with registry
# advice.
class FakeProjectWhoseLookupFails
  def VBComponents = self
  def Item(_name)
    raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800A0009)')
  end
end

class FakeParentForSheet
  def initialize(vb_project, denied: false)
    @vb_project = vb_project
    @denied = denied
  end

  def VBProject
    raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800A03EC)') if @denied

    @vb_project
  end
end

# Stands in for the COM Range that Sheet#[] and #[]= end up handing to
# WineOLE::MSOffice::Range -- just enough for Range#write's shape check.
class FakeComRangeForSheet
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

# Stands in for the COM Worksheet. Records what Range/Cells were asked for
# so tests can assert Sheet addressed the right cells without needing a
# real Excel.
class FakeComSheet
  attr_reader :range_calls, :cells_calls, :ranges

  def initialize(vb_project: nil, denied: false)
    @range_calls = []
    @cells_calls = []
    @ranges = []
    @vb_project = vb_project
    @denied = denied
  end

  def Range(addr)
    @range_calls << addr
    fake = FakeComRangeForSheet.new
    @ranges << fake
    fake
  end

  def Cells(row, col)
    @cells_calls << [row, col]
    fake = FakeComRangeForSheet.new
    @ranges << fake
    fake
  end

  def Name
    'Sheet1'
  end

  # Measured against a live Excel 11: Worksheet.CodeName is the module's
  # own name ("Sheet1"), independent of the visible tab Name.
  def CodeName
    'Sheet1'
  end

  # Measured against a live Excel 11: Worksheet.Parent is the Workbook.
  def Parent
    @parent ||= FakeParentForSheet.new(@vb_project, denied: @denied)
  end
end

class SheetTest < Minitest::Test
  def sheet(version: 11.0)
    WineOLE::MSOffice::Sheet.new(FakeComSheet.new, version: version)
  end

  # A sheet whose Parent.VBProject already has this sheet's own module
  # (CodeName "Sheet1") -- a worksheet's module cannot be created, only
  # found, so the fake must start with it already there.
  def sheet_with_project(version: 11.0)
    project = FakeVBProjectForSheet.new
    project.add_existing('Sheet1')
    WineOLE::MSOffice::Sheet.new(FakeComSheet.new(vb_project: project), version: version)
  end

  # A sheet whose Parent.VBProject raises the way Excel does when the
  # registry refuses VBA access.
  def sheet_denied(version: 11.0)
    WineOLE::MSOffice::Sheet.new(FakeComSheet.new(denied: true), version: version)
  end

  # --- [] with two integers -> Cells ---------------------------------

  def test_bracket_with_two_integers_wraps_cells
    s = sheet
    r = s[2, 3]
    assert_instance_of WineOLE::MSOffice::Range, r
    assert_equal [[2, 3]], s.ole.cells_calls
  end

  # --- [] with a string address -> Range ------------------------------

  def test_bracket_with_a_string_address_wraps_range
    s = sheet
    r = s['A1:B2']
    assert_instance_of WineOLE::MSOffice::Range, r
    assert_equal ['A1:B2'], s.ole.range_calls
  end

  # --- version is plumbed through to Address.parse --------------------

  def test_version_controls_which_grid_is_accepted
    # XFD1 is beyond Excel 11's IV/65536 grid, but within Excel 12's.
    err = assert_raises(ArgumentError) { sheet(version: 11.0)['XFD1'] }
    assert_match(/range/, err.message)

    r = sheet(version: 12.0)['XFD1']
    assert_instance_of WineOLE::MSOffice::Range, r
  end

  # --- [] = delegates to Range#write, never #fill ----------------------

  def test_bracket_assign_writes_through_range_write
    s = sheet
    s['A1:B1'] = 7
    assert_equal 7, s.ole.ranges.last.written
  end

  def test_bracket_assign_calls_write_not_fill
    s = sheet
    # write() raises on a flat array that does not exactly fit the range;
    # fill() would happily replicate/pad it instead and never raise. The
    # fake range always reports itself as 1x1, so any multi-element array
    # proves it is write(), not fill(), backing []=.
    err = assert_raises(ArgumentError) { s['A1:B1'] = [1, 2, 3] }
    assert_match(/1x1/, err.message)
    assert_match(/3 elements/, err.message)
  end

  def test_bracket_assign_with_two_integers
    s = sheet
    s[1, 1] = 5
    assert_equal [[1, 1]], s.ole.cells_calls
  end

  # --- addresses without a range raise, both for read and for write ----

  def test_assigning_to_an_address_without_a_range_raises
    s = sheet
    err = assert_raises(ArgumentError) { s[''] = 0 }
    assert_match(/range/, err.message)
  end

  def test_reading_an_address_without_a_range_also_raises
    s = sheet
    err = assert_raises(ArgumentError) { s[''] }
    assert_match(/range/, err.message)
  end

  # --- an address naming another sheet is refused, not silently honoured ---

  def test_an_address_naming_another_worksheet_raises
    s = sheet
    err = assert_raises(ArgumentError) { s['Sheet2!A1'] }
    assert_match(/sheet/i, err.message)
  end

  def test_an_address_naming_a_workbook_raises
    s = sheet
    err = assert_raises(ArgumentError) { s['[Book2]Sheet1!A1'] }
    assert_match(/sheet/i, err.message)
  end

  # --- passthrough ------------------------------------------------------

  def test_unknown_methods_go_to_com
    assert_equal 'Sheet1', sheet.Name
  end

  def test_ole_reader_exposes_the_underlying_proxy
    s = sheet
    assert_instance_of FakeComSheet, s.ole
  end

  # The denial rescue used to wrap the whole lookup, so ANY COM failure
  # after the project opened came back as "turn on AccessVBOM" -- advice
  # for a condition the caller is not in. Only the VBProject fetch is the
  # denial.
  def test_a_failure_after_the_project_opens_is_not_reported_as_access_denied
    s = WineOLE::MSOffice::Sheet.new(
      FakeComSheet.new(vb_project: FakeProjectWhoseLookupFails.new), version: 11.0
    )
    err = assert_raises(WineOLE::RemoteError) { s.vba.write('Sub A()', name: 'a') }
    assert_match(/0x800A0009/, err.message)
    refute_kind_of WineOLE::MSOffice::VBA::AccessDenied, err,
                   'a missing component is not a permissions problem'
  end

  # --- vba / remove_vba ---------------------------------------------------

  def test_sheet_vba_writes_into_this_sheets_own_module
    s = sheet_with_project
    s.vba.write("Private Sub Go_Click()\nEnd Sub", name: 'go')
    mod = s.ole.Parent.VBProject.components['Sheet1']
    assert_includes mod.CodeModule.text, "'<wineole:go>"
  end

  def test_sheet_remove_vba_leaves_the_module_in_place
    s = sheet_with_project
    s.vba.write('Sub Go()', name: 'go')
    s.vba.remove('go')
    assert s.ole.Parent.VBProject.components.key?('Sheet1'),
      "a sheet's module cannot be removed, and must not be attempted"
    refute_includes s.ole.Parent.VBProject.components['Sheet1'].CodeModule.text, 'Sub Go()'
  end

  # --- a denied VBProject gives advice, not a raw RemoteError ------------
  #
  # own_code_module reaches VBProject exactly the way Book#vb_project does,
  # for the same reason: 0x800A03EC and the localized message it comes
  # with cannot be told apart from a rejected NumberFormat, so the registry
  # is what turns the refusal into advice. Mirrors book_test.rb's
  # stub_vba_state -- Minitest::Test#stub does not exist on this host's
  # minitest 6.0.6, so the swap is done by hand.
  def stub_vba_state(result)
    original = WineOLE::MSOffice::VBA.method(:run_reg)
    WineOLE::MSOffice::VBA.define_singleton_method(:run_reg) { |*_args| result }
    yield
  ensure
    WineOLE::MSOffice::VBA.define_singleton_method(:run_reg, original)
  end

  def test_sheet_vba_on_a_denied_workbook_says_what_to_do
    s = sheet_denied
    err = nil
    stub_vba_state(['reg: <a localized not-found message>', false]) do
      err = assert_raises(WineOLE::MSOffice::VBA::AccessDenied) { s.vba.write('Sub A()', name: 'a') }
    end
    assert_match(/wineole-vba enable/, err.message)
    assert_match(/restart Excel/i, err.message)
  end

  def test_sheet_remove_vba_on_a_denied_workbook_says_what_to_do
    s = sheet_denied
    err = nil
    stub_vba_state(['reg: <a localized not-found message>', false]) do
      err = assert_raises(WineOLE::MSOffice::VBA::AccessDenied) { s.vba.remove('a') }
    end
    assert_match(/wineole-vba enable/, err.message)
    assert_match(/restart Excel/i, err.message)
  end

  # The registry-already-enabled branch: access is still refused because a
  # running Excel caches the setting from startup, so the message must
  # steer toward restarting Excel instead of re-running `wineole-vba
  # enable`.
  def test_sheet_vba_on_a_denied_workbook_with_the_registry_already_enabled_says_restart_excel
    s = sheet_denied
    err = nil
    stub_vba_state(["    AccessVBOM    REG_DWORD    0x1\r\n", true]) do
      err = assert_raises(WineOLE::MSOffice::VBA::AccessDenied) { s.vba.write('Sub A()', name: 'a') }
    end
    assert_match(/restart Excel/i, err.message)
    refute_match(/wineole-vba enable/, err.message,
      'the registry is already enabled -- telling the reader to enable it again is wrong')
  end
end
