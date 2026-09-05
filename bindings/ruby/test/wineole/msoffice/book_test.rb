require 'minitest/autorun'
require 'tmpdir'
require_relative '../../../lib/wineole/msoffice/book'

# Stands in for a WineOLE::Client, answering only what Paths.convertible?
# and Book need.
FakeClientForBook = Struct.new(:loopback) do
  def loopback?
    loopback
  end
end

# Stands in for the COM Workbook. Records what SaveAs received, and what
# Close received, so tests can assert on the exact values that would have
# crossed the wire -- without launching Excel.
class FakeComWorkbookForBook
  attr_reader :save_as_calls, :close_calls
  attr_accessor :path, :full_name, :vb_project_denied
  attr_writer :vb_project

  def initialize(path: '', full_name: '')
    @path = path
    @full_name = full_name
    @save_as_calls = []
    @close_calls = []
  end

  def VBProject
    raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800A03EC)') if @vb_project_denied

    @vb_project ||= FakeVBProject.new
  end

  def SaveAs(path)
    @save_as_calls << path
  end

  def Path
    @path
  end

  def FullName
    @full_name
  end

  def Close(save_changes)
    @close_calls << save_changes
  end

  # A distinct object, as real COM has it: a Workbook answers Worksheets,
  # and the *collection* answers Item and Count. Returning self here would
  # let a regression to @ole.Item(...) pass unnoticed.
  def Worksheets
    @worksheets ||= FakeComWorksheetsForBook.new
  end
end

class FakeComWorksheetsForBook
  def Item(name_or_index)
    FakeComWorksheetForBook.new(name_or_index)
  end

  def Count
    2
  end
end

class FakeComWorksheetForBook
  attr_reader :name_or_index

  def initialize(name_or_index)
    @name_or_index = name_or_index
  end
end

# Stands in for a VBA CodeModule. Holds text the way Excel does: reads come
# back with CRLF whatever was written, and AddFromString inserts at the top.
# Records how many times the whole body was fetched, so a round-trip count
# can be asserted (mirrors the fake in vba_block_test.rb -- kept in sync by
# hand so the two do not drift).
class FakeCodeModule
  attr_reader :reads

  # `lines:` bypasses the text split entirely, so a fake can hold blank
  # lines directly -- "\r\n".split(/\r?\n/) is [] in Ruby, not the two blank
  # lines Excel actually reports for a module emptied of its blocks.
  def initialize(text = '', lines: nil)
    @lines = lines || (text.empty? ? [] : text.split(/\r?\n/))
    @reads = 0
  end

  def CountOfLines = @lines.length

  def Lines(start, count)
    @reads += 1
    raise 'Lines(1, 0) must never be called' if count.zero?

    @lines[(start - 1), count].join("\r\n") + "\r\n"
  end

  def AddFromString(t) = @lines = t.split(/\r?\n/) + @lines
  def DeleteLines(start, count) = @lines.slice!(start - 1, count)
  def text = @lines.join("\n")
end

class FakeVBComponent
  # Name is writable because Excel's own Add() names the component and the
  # caller renames it afterwards -- the two steps are not atomic there
  # either, which is why Book#module_named only ever renames a fresh one.
  attr_accessor :export_bytes
  attr_reader :Name

  # Type mirrors VBComponents' own: 1 standard, 2 class, 3 UserForm, 100 a
  # module Excel owns (a worksheet's, ThisWorkbook's).
  attr_accessor :Type

  # Excel's Add-then-rename is not atomic: adding under a taken name
  # succeeds and only the rename fails, leaving a stray behind. The fake
  # has to be able to reach that state or the cleanup cannot be tested.
  attr_accessor :rename_fails

  def Name=(value)
    raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x80020009)') if @rename_fails

    @Name = value
  end

  def initialize(name, type: 1)
    @Name = name
    @Type = type
    @rename_fails = false
    @code_module = FakeCodeModule.new
    @export_bytes = "Attribute VB_Name = \"#{name}\"\r\n".b
  end

  def CodeModule = @code_module

  # Excel writes the file itself; the fake just puts the bytes there.
  #
  # book.rb hands this a Windows-shaped path when the (fake) bridge is
  # loopback -- real Wine resolves that back to the real file transparently
  # at the OS layer, which this reproduces with Paths.to_local so the fake
  # can still do plain Ruby File I/O on it. A no-op when the host has no
  # real winepath to have converted the path in the first place.
  def Export(path) = File.binwrite(WineOLE::MSOffice::Paths.to_local(path), @export_bytes)
end

class FakeVBComponents
  def initialize(project) = @project = project
  def Count = @project.component_list.length

  def Item(name)
    @project.component_list.find { |c| c.Name == name } ||
      raise(WineOLE::RemoteError.new('X', 'not found'))
  end

  def Add(type)
    c = FakeVBComponent.new("Module#{@project.component_list.length + 1}", type: type)
    c.rename_fails = @project.next_rename_fails
    @project.next_rename_fails = false
    @project.component_list << c
    c
  end

  def Remove(component) = @project.component_list.delete(component)

  # Same Windows-path-shaped-input reasoning as FakeVBComponent#Export above.
  def Import(path)
    @project.imported_bytes = File.binread(WineOLE::MSOffice::Paths.to_local(path))
    c = FakeVBComponent.new('Imported')
    @project.component_list << c
    c
  end
end

class FakeVBProject
  attr_reader :component_list
  attr_accessor :imported_bytes, :next_rename_fails

  def initialize
    @component_list = []
    @imported_bytes = nil
    @next_rename_fails = false
  end

  # A Hash view keyed by each component's *current* Name, built fresh on
  # every call. It cannot be a plain Hash keyed at Add() time: Add() and
  # Name= are two separate steps (see FakeVBComponent above and
  # Book#module_named), and a Hash key does not follow a rename -- Excel's
  # own VBComponents.Item looks up by whatever the Name currently is.
  def components
    @component_list.each_with_object({}) { |c, h| h[c.Name] = c }
  end

  def VBComponents = @vb_components ||= FakeVBComponents.new(self)

  def add_existing(name, type: 1)
    c = FakeVBComponent.new(name, type: type)
    @component_list << c
    c
  end
end

class BookTest < Minitest::Test
  def winepath?
    return @winepath if defined?(@winepath)

    @winepath = system('which winepath > /dev/null 2>&1')
  end

  def book(ole: FakeComWorkbookForBook.new, loopback: true, convert_paths: true, version: 11.0, vb_project: nil)
    case vb_project
    when :denied         then ole.vb_project_denied = true
    when FakeVBProject   then ole.vb_project = vb_project
    when nil             then nil
    else raise ArgumentError, "unsupported vb_project #{vb_project.inspect}"
    end
    WineOLE::MSOffice::Book.new(
      ole, client: FakeClientForBook.new(loopback), version: version, convert_paths: convert_paths
    )
  end

  # --- save_as: path conversion is gated correctly ----------------------

  def test_a_remote_clients_save_as_passes_the_path_unconverted
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole, loopback: false)
    b.save_as('/home/user/out.xls')
    assert_equal ['/home/user/out.xls'], ole.save_as_calls
  end

  def test_convert_paths_false_on_a_loopback_client_passes_the_path_unconverted
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole, loopback: true, convert_paths: false)
    b.save_as('/home/user/out.xls')
    assert_equal ['/home/user/out.xls'], ole.save_as_calls
  end

  def test_a_loopback_clients_save_as_hands_com_a_windows_shaped_path
    skip 'winepath is not on PATH' unless winepath?

    ole = FakeComWorkbookForBook.new
    b = book(ole: ole, loopback: true, convert_paths: true)
    b.save_as('/home/user/out.xls')
    assert_equal 1, ole.save_as_calls.length
    assert_match(/\A(?:[A-Za-z]:[\\\/]|\\\\)/, ole.save_as_calls.first,
      'a loopback client should get a Windows-shaped path back from winepath')
  end

  # --- convert_paths is false for a caller who cannot legitimately opt in ---

  def test_a_remote_client_cannot_be_talked_into_converting
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole, loopback: false, convert_paths: true)
    b.save_as('/home/user/out.xls')
    assert_equal ['/home/user/out.xls'], ole.save_as_calls,
      "a caller passing convert_paths: true must not override a remote bridge's " \
      'own answer about whether converting means anything'
  end

  # --- local_path -----------------------------------------------------

  def test_local_path_is_left_alone_when_not_converting
    ole = FakeComWorkbookForBook.new(path: 'Z:\\home\\user')
    b = book(ole: ole, loopback: false)
    assert_equal 'Z:\\home\\user', b.local_path
  end

  def test_local_path_of_an_unsaved_book_is_empty_without_shelling_out
    ole = FakeComWorkbookForBook.new(path: '')
    b = book(ole: ole, loopback: true, convert_paths: true)
    assert_equal '', b.local_path
  end

  # --- local_file -------------------------------------------------------
  #
  # Gated the same way local_path is -- unlike the README's own
  # Paths.to_local(book.FullName) recipe, which is ungated and, over a
  # remote bridge, would silently run a local winepath over a path that
  # names some other machine's filesystem.

  def test_a_remote_clients_local_file_is_left_unconverted
    ole = FakeComWorkbookForBook.new(full_name: 'Z:\\tmp\\out.xls')
    b = book(ole: ole, loopback: false)
    assert_equal 'Z:\\tmp\\out.xls', b.local_file
  end

  def test_a_loopback_clients_local_file_is_a_linux_path
    skip 'winepath is not on PATH' unless winepath?

    ole = FakeComWorkbookForBook.new(full_name: 'Z:\\tmp\\out.xls')
    b = book(ole: ole, loopback: true, convert_paths: true)
    refute_match(/\A(?:[A-Za-z]:[\\\/]|\\\\)/, b.local_file,
      'a loopback client should get a Linux path back from winepath')
  end

  def test_local_file_of_an_unsaved_book_does_not_shell_out
    ole = FakeComWorkbookForBook.new(full_name: 'Book1')
    b = book(ole: ole, loopback: true, convert_paths: true)
    assert_equal 'Book1', b.local_file
  end

  # --- sheet / each_sheet go through Worksheets, not Sheets -------------

  def test_sheet_wraps_a_worksheet_by_name_or_index
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole)
    s = b.sheet('Sheet1')
    assert_instance_of WineOLE::MSOffice::Sheet, s
    assert_equal 'Sheet1', s.ole.name_or_index
  end

  def test_each_sheet_without_a_block_returns_an_enumerator
    b = book
    enum = b.each_sheet
    assert_instance_of Enumerator, enum
    assert_equal 2, enum.to_a.length
    assert enum.to_a.all? { |s| s.is_a?(WineOLE::MSOffice::Sheet) }
  end

  def test_each_sheet_yields_each_sheet
    b = book
    seen = []
    b.each_sheet { |s| seen << s }
    assert_equal 2, seen.length
  end

  # --- close is a deliberate shadow, not the raw COM member -------------

  def test_close_defaults_to_not_saving
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole)
    b.close
    assert_equal [false], ole.close_calls
  end

  def test_close_can_be_told_to_save
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole)
    b.close(save: true)
    assert_equal [true], ole.close_calls
  end

  # --- the raw COM member stays reachable in PascalCase ------------------

  def test_the_raw_com_close_stays_reachable_in_pascal_case
    ole = FakeComWorkbookForBook.new
    b = book(ole: ole)
    b.Close(true)
    assert_equal [true], ole.close_calls
  end

  # --- vba / remove_vba ---------------------------------------------------

  def test_vba_puts_a_named_block_in_the_wrappers_own_module
    b = book
    b.vba.write("Sub Go()\nEnd Sub", name: 'go')
    mod = b.ole.VBProject.components['WineOLE']
    assert_includes mod.CodeModule.text, "'<wineole:go>"
    assert_includes mod.CodeModule.text, 'Sub Go()'
  end

  def test_removing_the_last_block_removes_the_module
    b = book
    b.vba.write('Sub Go()', name: 'go')
    b.vba.remove('go')
    refute b.ole.VBProject.components.key?('WineOLE'),
      "an empty module is litter -- it should go when its last block does"
  end

  def test_removing_one_of_two_blocks_keeps_the_module
    b = book
    b.vba.write('Sub A()', name: 'a')
    b.vba.write('Sub B()', name: 'b')
    b.vba.remove('a')
    assert b.ole.VBProject.components.key?('WineOLE')
    assert_includes b.ole.VBProject.components['WineOLE'].CodeModule.text, 'Sub B()'
  end

  # remove_vba asks VBABlock.remove to find and delete the block (one Lines
  # call) and must not fetch the body again afterwards just to ask whether
  # the module is now blank -- mirrors test_the_body_is_fetched_once_per_operation
  # in vba_block_test.rb, but for the whole remove_vba operation.
  # Leftover content matters here: removing the *only* block empties the
  # module down to CountOfLines == 0, and body() short-circuits on that
  # without a Lines call either way -- which would let this test pass by
  # coincidence even with the old double-fetch. Leaving a second block
  # behind means blank?'s body fetch (if remove_vba still made one) could
  # not be free, so this actually distinguishes the fix.
  def test_remove_vba_fetches_the_body_once
    b = book
    b.vba.write('Sub A()', name: 'a')
    b.vba.write('Sub B()', name: 'b')
    cm = b.ole.VBProject.components['WineOLE'].CodeModule
    reads_before = cm.reads

    b.vba.remove('a')

    assert_equal reads_before + 1, cm.reads, 'one Lines call, not one per operation'
  end

  # access_denied_message consults the real registry (VBA.state) to pick its
  # wording, so this test would otherwise pass or fail depending on whatever
  # AccessVBOM happens to be set to on the host running the suite -- as
  # observed here, where an earlier live-Excel session (Task 1's own
  # testing) had left it enabled. Isolated the same way vba_test.rb's
  # stub_reg isolates VBA.state: by swapping the singleton method, since
  # Minitest::Test#stub does not exist on this host's minitest 6.0.6.
  def stub_vba_state(result)
    original = WineOLE::MSOffice::VBA.method(:run_reg)
    WineOLE::MSOffice::VBA.define_singleton_method(:run_reg) { |*_args| result }
    yield
  ensure
    WineOLE::MSOffice::VBA.define_singleton_method(:run_reg, original)
  end

  def test_a_denied_project_says_what_to_do
    b = book(vb_project: :denied)
    err = nil
    stub_vba_state(['reg: <a localized not-found message>', false]) do
      err = assert_raises(WineOLE::MSOffice::VBA::AccessDenied) { b.vba.write('Sub A()', name: 'a') }
    end
    assert_match(/wineole-vba enable/, err.message)
    assert_match(/restart Excel/i, err.message)
  end

  # The spec's second denial row: the registry is already enabled and access
  # is still refused, because a running Excel caches the setting from
  # startup. That case cannot be told apart from the first by HRESULT or
  # message text, so access_denied_message must consult VBA.state itself --
  # this is the branch that only fires when state comes back :enabled.
  def test_a_denied_project_with_the_registry_already_enabled_says_restart_excel
    b = book(vb_project: :denied)
    err = nil
    stub_vba_state(["    AccessVBOM    REG_DWORD    0x1\r\n", true]) do
      err = assert_raises(WineOLE::MSOffice::VBA::AccessDenied) { b.vba.write('Sub A()', name: 'a') }
    end
    assert_match(/restart Excel/i, err.message)
    refute_match(/wineole-vba enable/, err.message,
      'the registry is already enabled -- telling the reader to enable it again is wrong')
  end

  # --- into: targets an existing component, not the wrapper's own -------

  def test_into_targets_an_existing_component
    b = book
    b.ole.VBProject.add_existing('AppForm')
    b.vba.write("Private Sub Go_Click()\nEnd Sub", name: 'go', into: 'AppForm')
    assert_includes b.ole.VBProject.components['AppForm'].CodeModule.text, "'<wineole:go>"
    refute b.ole.VBProject.components.key?('WineOLE'),
      'into: must not create the default module'
  end

  # We clean up what we made, not what we were pointed at. A UserForm can be
  # deleted, unlike ThisWorkbook -- so the rule has to be stated, not left to
  # what COM happens to allow.
  #
  # A same-named block is put in the wrapper's own module *first* -- without
  # it, a disabled `from:` branch falls through to the default path, finds no
  # "WineOLE" module at all, and does nothing whatsoever; AppForm would then
  # survive by accident rather than by a working guard, and the block would
  # never actually be removed from it either. Asserting that the block is
  # gone from AppForm -- the thing the call was for -- catches that; the
  # WineOLE module surviving untouched catches the fall-through itself, in
  # case it lands on something to damage instead of nothing at all.
  def test_a_named_component_is_never_removed_even_when_it_empties
    b = book
    b.vba.write('Sub Go()', name: 'go')
    b.ole.VBProject.add_existing('AppForm')
    b.vba.write('Sub Go()', name: 'go', into: 'AppForm')

    b.vba.remove('go', from: 'AppForm')

    assert b.ole.VBProject.components.key?('AppForm')
    refute_includes b.ole.VBProject.components['AppForm'].CodeModule.text, 'Sub Go()',
      'the block must actually be removed from the named component the call was for'
    assert b.ole.VBProject.components.key?('WineOLE'),
      "a call naming from: must never fall through and touch the wrapper's own module"
    assert_includes b.ole.VBProject.components['WineOLE'].CodeModule.text, "'<wineole:go>"
  end

  def test_into_a_component_that_is_not_there_says_so
    b = book
    err = assert_raises(ArgumentError) { b.vba.write('Sub A()', name: 'a', into: 'Missing') }
    assert_match(/Missing/, err.message)
  end

  # The from: side of the same guard -- named_component! is shared by vba
  # and remove_vba, but nothing had pinned that remove_vba actually raises
  # before touching anything, rather than e.g. silently removing from the
  # wrapper's own module.
  def test_remove_vba_from_a_component_that_is_not_there_says_so
    b = book
    err = assert_raises(ArgumentError) { b.vba.remove('a', from: 'Missing') }
    assert_match(/Missing/, err.message)
  end

  # --- import_vba / export_vba -------------------------------------------

  # minitest 6 has no #stub, so swap the singleton method by hand and put it
  # back in an ensure.
  def with_codepage(name)
    v = WineOLE::MSOffice::VBA
    original = v.method(:codepage)
    v.define_singleton_method(:codepage) { name }
    yield
  ensure
    v.define_singleton_method(:codepage, original)
  end

  def test_export_converts_the_codepage_and_the_line_endings
    b = book
    b.ole.VBProject.add_existing('Mod')
    # What Excel writes: the ANSI codepage, CRLF.
    b.ole.VBProject.components['Mod'].export_bytes =
      "Attribute VB_Name = \"Mod\"\r\n' caf\xE9\r\n".b

    Dir.mktmpdir do |dir|
      out = File.join(dir, 'Mod.bas')
      with_codepage('CP1252') { b.vba.export('Mod', out) }
      text = File.binread(out)
      assert_equal "Attribute VB_Name = \"Mod\"\n' café\n", text.force_encoding('UTF-8')
      refute_includes text, "\r", 'a file written to a Linux path should not carry CRLF'
    end
  end

  def test_import_converts_utf8_into_the_codepage
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      File.write(src, "Attribute VB_Name = \"Mod\"\n' café\n")
      with_codepage('CP1252') { b.vba.import(src) }
      handed = b.ole.VBProject.imported_bytes
      assert_equal "' caf\xE9".b, handed[/' caf./m].b,
        'Excel must be handed the codepage bytes, not UTF-8'
    end
  end

  # Silently substituting would manufacture the very failure this wrapper
  # exists to avoid: it succeeded, and the result is wrong.
  def test_a_character_the_codepage_cannot_hold_raises
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      File.write(src, "' \u{1F600}\n")
      # This used to assert the bare Encoding::UndefinedConversionError,
      # which pinned the defect rather than the behaviour: that exception
      # names neither the file nor the way out, and the string path gave a
      # full explanation for the identical condition.
      err = assert_raises(ArgumentError) do
        with_codepage('CP1252') { b.vba.import(src) }
      end
      assert_match(/"\u{1F600}"/, err.message, 'the character that stopped it')
      assert_match(/CP1252/, err.message)
    end
  end

  # The file is handed to Excel by path. On a remote bridge that path means
  # a different machine's filesystem, so there is nothing sensible to do.
  def test_import_and_export_refuse_a_remote_bridge
    b = book(loopback: false)
    assert_raises(ArgumentError) { b.vba.import('/tmp/x.bas') }
    assert_raises(ArgumentError) { b.vba.export('Mod', '/tmp/x.bas') }
  end

  # The write direction had been left bare: a file that DECODES cleanly can
  # still hold a character the codepage cannot store, and #encode raised a
  # naked Encoding::UndefinedConversionError naming neither the file nor
  # the way out -- while the string path gave a full explanation for the
  # very same condition. One message now serves both.
  def test_a_character_the_codepage_cannot_store_is_refused_by_name_on_the_file_path
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      File.binwrite(src, "' caf\u00e9 \u2713\n".encode('UTF-8'))
      err = assert_raises(ArgumentError) { with_codepage('CP932') { b.vba.import(src) } }
      assert_match(/CP932/, err.message)
      assert_match(/#{Regexp.escape(src)}/, err.message, 'the message must name the file')
      assert_match(/ChrW/, err.message, 'and the way out')
    end
  end

  # Both paths are bound by the same codepage, so a caller must not be able
  # to tell from the message which one they hit -- only which character
  # stopped it. Two texts, one explanation.
  def test_the_string_path_and_the_file_path_explain_the_same_thing
    b = book
    from_string = assert_raises(ArgumentError) do
      with_codepage('CP932') { b.vba.write("x = \"caf\u00e9\"", name: 'a') }
    end
    from_file = nil
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      File.binwrite(src, "x = \"caf\u00e9\"\n".encode('UTF-8'))
      from_file = assert_raises(ArgumentError) { with_codepage('CP932') { b.vba.import(src) } }
    end
    tail = ->(m) { m[/which the system codepage.*/m] }
    assert_equal tail.call(from_string.message), tail.call(from_file.message)
  end

  # --- add_component / remove_component ---------------------------------

  def test_add_component_creates_one_of_the_named_kind
    b = book(vb_project: FakeVBProject.new)
    c = b.vba.add_component('Utils')
    assert_equal 'Utils', c.Name
    assert_equal 1, c.Type, 'standard module by default'
    assert_equal 3, b.vba.add_component('Dialog', kind: :form).Type
    assert_equal 2, b.vba.add_component('Thing', kind: :class).Type
  end

  # The name check has to happen BEFORE Add, not after. Excel's Add and the
  # rename that follows it are not atomic -- adding under a taken name
  # succeeds and only the rename fails, leaving a stray Module1 behind. So
  # the assertion that matters is not just "it raised" but "it added
  # nothing": a raise after a successful Add would still leave the litter.
  def test_add_component_refuses_a_taken_name_without_adding_anything
    project = FakeVBProject.new
    project.add_existing('Utils')
    b = book(vb_project: project)

    err = assert_raises(ArgumentError) { b.vba.add_component('Utils') }
    assert_match(/already has a VBA component named "Utils"/, err.message)
    assert_equal 1, project.component_list.length, 'nothing may be added on the refusal path'
  end

  def test_add_component_refuses_an_unknown_kind_without_adding_anything
    project = FakeVBProject.new
    b = book(vb_project: project)

    err = assert_raises(ArgumentError) { b.vba.add_component('Utils', kind: :worksheet) }
    assert_match(/unknown component kind :worksheet/, err.message)
    assert_empty project.component_list
  end

  # The pre-check cannot cover a name Excel itself rejects (too long, an
  # illegal character), so the Add can still succeed with the rename
  # failing after it. The stray has to be taken back out.
  def test_a_rename_that_fails_after_add_takes_the_stray_back_out
    project = FakeVBProject.new
    project.next_rename_fails = true
    b = book(vb_project: project)

    err = assert_raises(ArgumentError) { b.vba.add_component('this name is refused') }
    assert_match(/Excel refused/, err.message)
    assert_empty project.component_list, 'the half-made component must not survive the failure'
  end

  def test_remove_component_deletes_it
    project = FakeVBProject.new
    project.add_existing('Utils')
    b = book(vb_project: project)

    b.vba.remove_component('Utils')
    assert_empty project.component_list
  end

  # A worksheet's module and ThisWorkbook are Excel's, not ours. COM
  # refuses to remove them, and a raw refusal says nothing useful, so this
  # is caught before the call and answered with what to do instead.
  def test_remove_component_refuses_a_module_excel_owns
    project = FakeVBProject.new
    project.add_existing('ThisWorkbook', type: 100)
    b = book(vb_project: project)

    err = assert_raises(ArgumentError) { b.vba.remove_component('ThisWorkbook') }
    assert_match(/cannot be deleted/, err.message)
    assert_match(/remove\(name, from:\)/, err.message, 'and must point at what does work')
    assert_equal 1, project.component_list.length, 'and must not have removed it'
  end

  def test_remove_component_on_a_name_that_is_not_there_says_so
    b = book(vb_project: FakeVBProject.new)
    err = assert_raises(ArgumentError) { b.vba.remove_component('Nope') }
    assert_match(/no VBA component named "Nope"/, err.message)
  end

  # --- import decides the encoding on evidence, never on a guess ---------

  # The behaviour this replaced: a file already in the codepage used to be
  # REFUSED, on the reasoning that it was probably a mistake. It is not a
  # mistake -- Excel's own Export writes the codepage, so every .bas from a
  # Windows toolchain arrives this way, and refusing them made the most
  # ordinary file in the ecosystem unusable.
  #
  # Bytes that are not valid UTF-8 PROVE the file is not UTF-8, and that is
  # evidence, not a guess. So they are read as the codepage now.
  def test_a_codepage_file_is_read_as_the_codepage_not_refused
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      # 0x92 is CP1252's curly apostrophe and is not valid UTF-8 in any
      # position, so UTF-8 is ruled out on evidence.
      File.binwrite(src, "' caf\x92\n".b)
      with_codepage('CP1252') { b.vba.import(src) }
      # Handed to Excel unchanged: it was already in the codepage Excel
      # wants, so the round trip through UTF-8 has to be lossless.
      assert_equal "' caf\x92\n".b, b.ole.VBProject.imported_bytes
    end
  end

  # A BOM is the one conclusive signal, so it wins over everything else --
  # including bytes that would otherwise be read as the codepage.
  def test_a_bom_decides_the_encoding_and_is_not_passed_through
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      File.binwrite(src, "\xEF\xBB\xBF' caf\xC3\xA9\n".b)
      with_codepage('CP1252') { b.vba.import(src) }
      handed = b.ole.VBProject.imported_bytes
      assert_equal "' caf\xE9\n".b, handed, 'decoded as UTF-8 and re-encoded to CP1252'
      refute handed.start_with?("\xEF\xBB\xBF".b), 'the BOM must not reach the module text'
    end
  end

  # Both encodings wrong at once. UTF-8 is ruled out by the bytes, and the
  # codepage cannot read them either, so there is genuinely nothing to
  # infer -- and the error has to say that rather than surfacing a bare
  # Encoding::InvalidByteSequenceError that names neither the file nor why
  # that encoding was the one attempted.
  def test_bytes_valid_in_neither_encoding_say_so_instead_of_raising_a_bare_encoding_error
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      # 0x92 leads a CP932 double-byte character; a bare \n cannot follow it.
      File.binwrite(src, "' caf\x92\n".b)
      err = assert_raises(ArgumentError) { with_codepage('CP932') { b.vba.import(src) } }
      assert_match(/#{Regexp.escape(src)}/, err.message, 'the message must name the file')
      assert_match(/CP932/, err.message, 'and the encoding it tried')
      assert_match(/not valid UTF-8/, err.message, 'and why that encoding was the one tried')
    end
  end

  # An explicit encoding: skips detection entirely -- and when it is wrong,
  # the message must not blame a detection that never ran.
  def test_an_explicit_encoding_skips_detection_and_owns_the_failure
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      File.binwrite(src, "' caf\x92\n".b)
      err = assert_raises(ArgumentError) { b.vba.import(src, encoding: 'CP932') }
      assert_match(/you passed encoding/, err.message)
      refute_match(/no BOM/, err.message, 'detection did not run, so it must not be cited')
    end
  end

  def test_a_codepage_byte_pair_that_happens_to_be_valid_utf8_is_not_caught
    b = book
    Dir.mktmpdir do |dir|
      src = File.join(dir, 'Mod.bas')
      # \xC2\xA0 is CP1252's non-breaking space glued to itself becoming
      # U+00A0 when read as UTF-8 -- the file is genuinely valid UTF-8 by
      # the only test available (valid_encoding?), so import_vba proceeds.
      File.binwrite(src, "' \xC2\xA0\n".b)
      with_codepage('CP1252') { b.vba.import(src) }
      handed = b.ole.VBProject.imported_bytes
      # Round-tripped as a single CP1252 byte (0xA0), not the two original
      # bytes -- this is the silent content change the Important finding
      # describes, left in place because it cannot be detected.
      assert_equal "' \xA0\n".b, handed
    end
  end

  # --- export_vba's raise-don't-substitute side ---------------------------

  # Symmetric to test_a_character_the_codepage_cannot_hold_raises above:
  # Excel's exported bytes can contain something undefined in its own
  # reported codepage (0x81 has no character in CP1252), and
  # force_encoding(cp).encode('UTF-8') must raise rather than substitute.
  #
  # This is Encoding::UndefinedConversionError, not InvalidByteSequenceError
  # -- checked directly against Ruby rather than assumed: CP1252 is a
  # single-byte encoding, so there is no such thing as an invalid *byte
  # sequence* in it, only bytes with no mapping to the target encoding,
  # which is exactly what UndefinedConversionError means.
  def test_export_raises_when_the_reported_codepage_cannot_decode_the_bytes
    b = book
    b.ole.VBProject.add_existing('Mod')
    b.ole.VBProject.components['Mod'].export_bytes =
      "Attribute VB_Name = \"Mod\"\r\n' \x81\r\n".b

    Dir.mktmpdir do |dir|
      out = File.join(dir, 'Mod.bas')
      assert_raises(Encoding::UndefinedConversionError) do
        with_codepage('CP1252') { b.vba.export('Mod', out) }
      end
    end
  end

  # --- a path ending in ".." gets a clear error, not Errno::EISDIR -------

  def test_import_a_path_ending_in_dotdot_says_so_clearly
    b = book
    Dir.mktmpdir do |dir|
      err = assert_raises(ArgumentError) { b.vba.import(File.join(dir, '..')) }
      assert_match(/\.\./, err.message)
    end
  end

  def test_export_a_path_ending_in_dotdot_says_so_clearly
    b = book
    b.ole.VBProject.add_existing('Mod')
    Dir.mktmpdir do |dir|
      err = assert_raises(ArgumentError) { b.vba.export('Mod', File.join(dir, '..')) }
      assert_match(/\.\./, err.message)
    end
  end
end
