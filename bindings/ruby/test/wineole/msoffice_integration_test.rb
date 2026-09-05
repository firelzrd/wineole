require 'minitest/autorun'
require 'tmpdir'
require_relative '../../lib/wineole/msoffice'
require_relative '../support/excel_integration_helper'

# Pins the assembled Office wrapper (Address/Paths/Range/Sheet/Book/Excel)
# against real Excel, running under Wine. Tasks 1-5 unit-tested every piece
# against fakes; this file is the first time the whole thing is exercised
# end to end, per spec §7's success criteria.
#
# Shares its spawn/teardown plumbing with wineole_integration_test.rb via
# ExcelIntegrationHelper -- see that file for why getting this dance wrong
# is exactly how a stray EXCEL.EXE survives a run.
class WineOLEMSOfficeIntegrationTest < Minitest::Test
  include ExcelIntegrationHelper

  # spec §7 criterion 1: Excel.run quits only what it created.
  #
  # Two halves, both against the same running bridge: run(:create) must
  # both create an instance and quit it again on the way out; a *separate*
  # instance started outside the wrapper, then merely attached to via
  # run(:connect_or_create), must survive the block untouched.
  def test_run_quits_only_what_it_created
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        assert xl.ole_created?, 'Excel.run(:create) must report the instance as newly created'
        xl.hide
        xl.no_alert { xl.ole.Workbooks.Add }
      end

      # Best-effort, not asserted: the same "over-release/late-release"
      # Wine COM flake documented in wineole_integration_test.rb (search
      # that file for IRemUnknown_RemRelease) shows up here too -- Quit()
      # itself is issued (run's ensure calling ole_release is what tells the
      # bridge to run CLEANUP_STEPS -- DisplayAlerts=false then Quit -- and
      # the msoffice unit tests pin that ole_release call with a fake that
      # records it), yet under Wine the *process* sometimes lingers for well
      # over 20s before actually exiting on its own, with no dialog and
      # nothing this wrapper controls to speed it up. Measured here: with no
      # workbook open, with an unsaved one left open, and with the workbook
      # explicitly closed first, the process still took more than 20s (but
      # eventually exited on its own, with no force-kill) in at least one of
      # those cases -- so asserting on it would make this test flake on
      # Wine's own timing, not on this wrapper's behavior.
      warn 'note: EXCEL.EXE outlived Excel.run(:create)\'s Quit for over 20s ' \
        '(known Wine COM release flake)' unless excel_gone?(timeout: 20)

      # A second instance, started by hand -- run(:connect_or_create) below
      # must attach to this one, not create a third, and must leave it
      # running afterwards.
      xl_raw = client.create('Excel.Application')
      xl_raw.Visible = false
      xl_raw.DisplayAlerts = false
      xl_raw.Workbooks.Add
      begin
        WineOLE::MSOffice::Excel.run(:connect_or_create, client: client) do |xl|
          refute xl.ole_created?,
            'Excel.run(:connect_or_create) must attach to the already-running instance, not create one'
        end

        # Still alive: a live Version call is proof, and doesn't race the
        # way polling a leftover-process list would.
        assert_equal '11.0', xl_raw.Version,
          'attaching via run(:connect_or_create) must not Quit the instance it did not create'
      ensure
        begin
          xl_raw.Quit
        rescue StandardError
          nil
        end
      end
    end
  end

  # Task 9: `leave_open` revokes the bridge's permission to run CLEANUP_STEPS
  # for this record at all -- not just for the ole_release run's ensure
  # issues, but past the bridge connection itself closing (with_bridge's own
  # ensure calls WineOLE.close right after this block). Asserted outside
  # with_bridge for exactly that reason: this is only interesting proof if
  # the Excel is still there once the connection that created it is gone.
  def test_leave_open_keeps_excel_running_after_run_and_after_the_connection_closes
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert { xl.ole.Workbooks.Add }
        xl.leave_open
      end
    end

    refute_empty excel_pids - @pre_existing_excel_pids,
      'leave_open must keep the Excel this run created alive past the block and the closed connection'
    # Left running deliberately -- teardown (ExcelIntegrationHelper) kills
    # any PID not in @pre_existing_excel_pids, so this does not leak.
  end

  # spec §7 criteria 2, 5, 6, 7: the addressing DSL entry point, the
  # always-2-D to_a, Excel's own bare-scalar Value, and passthrough to an
  # unwrapped COM member.
  def test_addressing_dsl_value_shapes_and_passthrough
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert do
          assert_equal '11.0', xl.version, 'xl.version must fall through to COM Application.Version'

          # New book + new sheet + cell, in one address.
          xl['[:new]:new!A1'] = 'hello'

          # ':last!' (bare worksheet, no book part) reaches the sheet just
          # created -- Worksheets.Add leaves it both the active sheet and
          # the last one.
          sheet = xl[':last!']
          assert_instance_of WineOLE::MSOffice::Sheet, sheet

          assert_equal [['hello']], sheet['A1'].to_a,
            'Range#to_a must always be 2-D, even for a single cell'
          assert_equal 'hello', sheet['A1'].Value,
            "Excel's own Value for a 1x1 range is a bare scalar -- to_a is what normalizes it"

          # Passthrough to a COM member this wrapper never defines.
          sheet.ole.PageSetup.Orientation = 2

          book = xl['[]']
          assert_instance_of WineOLE::MSOffice::Book, book
        end
      end
    end
  end

  # spec §7 criterion 3, the most valuable test in this task: the wrapper's
  # write lays a flat Array down a column correctly; Excel's own
  # Range.Value= assignment on the identical range replicates the first
  # element into every cell instead. Both are asserted here, side by side,
  # so the difference this wrapper exists to fix is visible in the test
  # itself.
  def test_write_lays_a_flat_array_down_the_column_unlike_excels_own_assignment
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert do
          xl.ole.Workbooks.Add
          sheet = xl[':last!']

          sheet['A1:A10'] = (1..10).to_a
          wrapped = sheet['A1:A10'].to_a.flatten
          assert_equal [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0], wrapped,
            'Range#write must lay a flat Array down a column in order'

          # The identical assignment, done Excel's own way (bypassing
          # write/shaped entirely via the raw COM Range).
          sheet.ole.Range('B1:B10').Value = (1..10).to_a
          raw = sheet.ole.Range('B1:B10').Value.flatten
          assert_equal Array.new(10, 1.0), raw,
            "Excel's own Range.Value= replicates a flat Array's first element down every " \
            'cell instead of laying it out -- this is exactly what Range#write exists to fix'

          refute_equal wrapped, raw,
            'the whole point of write: the wrapper and the raw COM assignment must disagree here'
        end
      end
    end
  end

  # spec §7 criterion 4: a value that does not fit the range's shape is a
  # dimension-mismatch exception, not a silent corruption.
  def test_write_raises_a_dimension_mismatch_error
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert do
          xl.ole.Workbooks.Add
          sheet = xl[':last!']

          err = assert_raises(ArgumentError) { sheet['E1:G3'] = [1, 2] }
          assert_match(/\Arange is 3x3;/, err.message)
          assert_match(/only fits a single row or column/, err.message)
        end
      end
    end
  end

  # spec §7 criterion 8: save_as takes a Linux path, and local_path hands
  # one back -- both ends of the Wine<->Linux path conversion Paths adds.
  def test_save_as_and_local_path_use_linux_paths
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert do
          xl['[:new]:new!A1'] = 'save-as check'
          book = xl['[]']

          Dir.mktmpdir('wineole-office-integration') do |dir|
            out_path = File.join(dir, 'out.xls')
            book.save_as(out_path)
            assert File.exist?(out_path),
              "save_as(#{out_path.inspect}) must create the file at that Linux path"

            assert_equal dir, book.local_path,
              'local_path must report the containing folder in Linux form, not a Wine one'

            book.close(save: false)
          end
        end
      end
    end
  end

  # The wrapper's colour vocabulary is RGB; Excel's is BGR. A round trip
  # through the number cannot catch a reversed conversion -- it would agree
  # with itself. Excel's own ColorIndex is what names the colour: 3 is red,
  # 5 is blue.
  def test_red_is_red_and_not_blue
    with_excel do |xl|
      sheet = xl[':first!']
      sheet['A1'].format(background: '#FF0000')
      assert_equal 3, sheet['A1'].ole.Interior.ColorIndex,
        'ColorIndex 3 is red; 5 would mean the RGB/BGR conversion is reversed'

      sheet['A2'].format(background: '#0000FF')
      assert_equal 5, sheet['A2'].ole.Interior.ColorIndex
    end
  end

  def test_format_reaches_a_whole_range_in_one_call
    with_excel do |xl|
      sheet = xl[':first!']
      sheet['C1:E3'].format(bold: true, align: :center)
      [sheet['C1'], sheet['E3']].each do |cell|
        assert_equal true, cell.ole.Font.Bold
        assert_equal(-4108, cell.ole.HorizontalAlignment)
      end
    end
  end

  # Color cannot express "no fill": a cleared cell and a white-painted cell
  # both report Color 16777215. Asserting on Color alone would accept an
  # implementation that paints white, so this compares every Interior
  # property against a cell that was never touched.
  def test_background_false_really_clears_rather_than_painting_white
    with_excel do |xl|
      sheet = xl[':first!']
      pristine = interior_snapshot(sheet['Z50'])

      cell = sheet['G1']
      cell.format(background: '#FF0000')
      refute_equal pristine, interior_snapshot(cell), 'the fill must actually have happened'

      cell.format(background: false)
      assert_equal pristine, interior_snapshot(cell),
        'a cleared cell must be indistinguishable from one that was never filled'
    end
  end

  def test_number_format_general_resets_whatever_the_locale
    with_excel do |xl|
      sheet = xl[':first!']
      pristine = sheet['Z50'].ole.NumberFormat

      cell = sheet['G3']
      cell.format(number_format: '0.00')
      refute_equal pristine, cell.ole.NumberFormat

      cell.format(number_format: :general)
      assert_equal pristine, cell.ole.NumberFormat
    end
  end

  # The reason :general exists at all. If this ever starts passing, the
  # wrapper is working around something that no longer happens.
  def test_writing_the_string_general_straight_to_com_still_fails
    with_excel do |xl|
      sheet = xl[':first!']
      assert_raises(WineOLE::RemoteError) do
        sheet['G4'].ole.NumberFormat = 'General'
      end
    end
  end

  def test_borders_reach_the_edges_they_name
    with_excel do |xl|
      sheet = xl[':first!']
      sheet['I1:K3'].format(border: :outline)
      assert_equal 1, sheet['I1:K3'].ole.Borders.Item(7).LineStyle
      assert_equal(-4142, sheet['I1:K3'].ole.Borders.Item(11).LineStyle,
        ':outline must leave the inside edges alone')
    end
  end

  # Format.write_border has a bulk path -- an assignment straight to the
  # Borders collection, used for :all, an explicit list of all six edges,
  # and border: false -- that the :outline test above never exercises,
  # because :outline is deliberately kept off that path (it would reach the
  # inside edges too). The bulk path's danger is different: Excel's
  # Borders collection also holds xlDiagonalDown (index 5) and xlDiagonalUp
  # (6) alongside the six edges this wrapper knows about. If a bulk
  # `Borders.LineStyle =` ever reached those, `format(border: :all)` would
  # silently draw an X through every cell in the range. Measured against a
  # live Excel 11 on a multi-cell range: both diagonals report -4142
  # (untouched) whether the range has never been formatted or was just
  # formatted with border: :all.
  def test_border_all_and_false_use_the_bulk_path_without_touching_the_diagonals
    with_excel do |xl|
      sheet = xl[':first!']
      range = sheet['I1:K3']
      borders = range.ole.Borders

      (5..12).each do |index|
        assert_equal(-4142, borders.Item(index).LineStyle, "edge #{index} must start untouched")
      end

      range.format(border: :all)
      assert_equal(-4142, borders.Item(5).LineStyle, 'border: :all must leave xlDiagonalDown alone')
      assert_equal(-4142, borders.Item(6).LineStyle, 'border: :all must leave xlDiagonalUp alone')
      (7..12).each do |index|
        assert_equal 1, borders.Item(index).LineStyle, "border: :all must set edge #{index}"
      end

      range.format(border: false)
      (5..12).each do |index|
        assert_equal(-4142, borders.Item(index).LineStyle, "border: false must clear edge #{index}")
      end
    end
  end

  # Excel 2003 has a 56-colour palette and silently approximates anything
  # else. Pinned here so the README's warning stays true rather than
  # becoming folklore.
  def test_excel_2003_snaps_an_off_palette_colour
    with_excel do |xl|
      sheet = xl[':first!']
      sheet['M1'].format(background: '#EEEEEE')
      assert_equal WineOLE::MSOffice::Color['#FFFFFF'],
        sheet['M1'].ole.Interior.Color.to_i,
        'Excel 2003 approximates #EEEEEE to pure white -- not a wrapper defect'
    end
  end

  # spec §7 criteria 1, 2, 5, 6, 7, 8: inject a named block, call it through
  # xl.Run, and see the definition survive.
  def test_a_named_block_survives_a_save_and_reopen
    with_excel do |xl|
      book = xl['[]']
      book.vba.write("Function Doubled(a)\n  Doubled = a * 2\nEnd Function", name: 'helpers')
      assert_equal 42, xl.ole.Run('Doubled', 21)
    end
  end

  # spec §7 criterion 2: rewriting a block must not leave the first
  # definition behind alongside the second.
  def test_rewriting_a_block_does_not_define_it_twice
    with_excel do |xl|
      book = xl['[]']
      book.vba.write("Function F(a)\n  F = a + 1\nEnd Function", name: 'f')
      book.vba.write("Function F(a)\n  F = a + 2\nEnd Function", name: 'f')
      assert_equal 23, xl.ole.Run('F', 21),
        'the second definition must win, and the first must be gone'
    end
  end

  # spec §7 criterion 4: the last block leaving a module removes the module
  # itself, not just the block's text.
  def test_removing_the_last_block_removes_the_module
    with_excel do |xl|
      book = xl['[]']
      book.vba.write('Function G()\n  G = 1\nEnd Function', name: 'g')
      before = book.ole.VBProject.VBComponents.Count
      book.vba.remove('g')
      assert_equal before - 1, book.ole.VBProject.VBComponents.Count
    end
  end

  # spec §7: a sheet's block must land in that sheet's own code module, not
  # the wrapper's default one -- that is where Excel looks for
  # <ActiveX control>_Click.
  def test_a_sheet_block_lands_in_that_sheets_module
    with_excel do |xl|
      sheet = xl[':first!']
      sheet.vba.write("Sub Marker()\nEnd Sub", name: 'marker')
      code_name = sheet.ole.CodeName
      body = xl['[]'].ole.VBProject.VBComponents.Item(code_name).CodeModule
      assert_includes body.Lines(1, body.CountOfLines), 'Sub Marker()'
    end
  end

  # Non-ASCII goes through COM as Unicode -- the bridge's own BSTR
  # marshalling (wineole-bridge/src/value.rs) never touches a codepage. But
  # this Excel's VBA6 engine stores module *source* per the local ANSI
  # codepage regardless of how the string arrived: text inside that
  # codepage's repertoire survives; text outside it does not raise, it is
  # just quietly downgraded. On this (CP932) host, Greek and Cyrillic
  # letters (JIS X 0208 rows 6-7) are inside the codepage and round-trip
  # intact -- measured
  # directly, ad hoc, before writing this test.
  def test_non_ascii_within_the_local_codepage_survives_string_injection
    with_excel do |xl|
      book = xl['[]']
      book.vba.write(%(Function Greet()\n  Greet = "αβγ-Дж"\nEnd Function), name: 'g')
      assert_equal 'αβγ-Дж', xl.ole.Run('Greet')
    end
  end

  # A module's text is held in the system ANSI codepage, so a character the
  # codepage cannot represent used to be substituted on the way in --
  # measured here, "café ✓" came back "cafe ?". An earlier measurement using
  # Japanese alone concluded this path carried Unicode, which was wrong:
  # CP932 simply represents Japanese.
  #
  # It is refused now, under the same rule import_vba follows. Silently
  # dropping an accent is the failure this phase exists to remove, and there
  # is no way to inject such a character as a literal at all -- so saying so
  # is the only honest option.
  def test_a_character_the_codepage_cannot_hold_is_refused_not_mangled
    with_excel do |xl|
      err = assert_raises(ArgumentError) do
        xl['[]'].vba.write(%(Function Greet()\n  Greet = "café ✓"\nEnd Function), name: 'g')
      end
      assert_match(/cannot represent/, err.message)
      assert_match(/é/, err.message, 'the message must name the character that stopped it')
    end
  end

  # spec §7 criteria 5, 6: export writes UTF-8/LF; import reads UTF-8 back
  # in, through the codepage boundary, and the result is still callable.
  def test_export_and_import_round_trip_through_utf8
    with_excel do |xl|
      book = xl['[]']
      book.vba.write("Function RT(a)\n  RT = a + 5\nEnd Function", name: 'rt')
      Dir.mktmpdir do |dir|
        out = File.join(dir, 'WineOLE.bas')
        book.vba.export('WineOLE', out)
        text = File.binread(out)
        assert_equal text, text.dup.force_encoding('UTF-8').encode('UTF-8'),
          'the exported file must be valid UTF-8'
        refute_includes text, "\r", 'and must use LF'

        book.vba.remove('rt')
        book.vba.import(out)
        assert_equal 26, xl.ole.Run('RT', 21)
      end
    end
  end

  # spec §7 criterion 5, with an actual non-ASCII payload crossing the
  # codepage boundary -- the ASCII-only round trip above proves the
  # mechanics (UTF-8/LF, importable) but never puts a non-ASCII byte
  # through the encode/decode step it exists to protect. A Greek/Cyrillic
  # literal is inside CP932, so (per the two tests above) it survives the
  # journey either way; this pins that export/import specifically does not
  # add its own corruption on top.
  def test_export_and_import_round_trip_preserves_non_ascii_within_the_codepage
    with_excel do |xl|
      book = xl['[]']
      book.vba.write(%(Function Payload()\n  Payload = "αβγ-Дж"\nEnd Function), name: 'k')
      Dir.mktmpdir do |dir|
        out = File.join(dir, 'WineOLE.bas')
        book.vba.export('WineOLE', out)
        assert_includes File.read(out, encoding: 'UTF-8'), 'αβγ-Дж'

        book.vba.remove('k')
        book.vba.import(out)
        assert_equal 'αβγ-Дж', xl.ole.Run('Payload')
      end
    end
  end

  # spec §7 criterion 8. This host has AccessVBOM enabled, so every test
  # above only exercises the allowed path -- the one path most users will
  # NOT hit first. This is the only test in the suite that exercises the
  # refusal, by switching the setting off, starting a fresh Excel (the
  # setting is read at startup, so a running instance would not notice),
  # and checking the refusal is the guidance message rather than a raw COM
  # error -- then restoring the setting whatever happens.
  def test_a_named_module_can_be_made_written_to_and_removed
    with_excel do |xl|
      book = xl['[]']
      book.vba.add_component('Utils')
      book.vba.write("Public Function Beta()\n  Beta = 8\nEnd Function", name: 'b', into: 'Utils')
      assert_equal 8, xl.ole.Run('Beta')

      book.vba.remove_component('Utils')
      # The component is gone, so the code in it has to be gone too --
      # asserting only that Remove returned would pass even if it had not.
      assert_raises(WineOLE::RemoteError) { xl.ole.Run('Beta') }
    end
  end

  # Excel Add()s before it renames, and the two are not atomic, so a
  # refusal that came after the Add would leave a stray Module1 behind.
  # Counting is what catches that; the raise alone would not.
  def test_making_a_component_under_a_taken_name_leaves_no_stray
    with_excel do |xl|
      book = xl['[]']
      book.vba.add_component('Utils')
      before = book.vba.project.VBComponents.Count

      assert_raises(ArgumentError) { book.vba.add_component('Utils') }
      assert_equal before, book.vba.project.VBComponents.Count
    end
  end

  def test_a_module_excel_owns_cannot_be_removed
    with_excel do |xl|
      err = assert_raises(ArgumentError) { xl['[]'].vba.remove_component('ThisWorkbook') }
      assert_match(/cannot be deleted/, err.message)
    end
  end

  # Where code lives decides whether it can be called at all -- the table
  # in BookVBA's comment, pinned against a live Excel so it cannot drift.
  def test_only_a_standard_module_answers_an_unqualified_run
    with_excel do |xl|
      book = xl['[]']
      book.vba.write("Public Function Alpha()\n  Alpha = 7\nEnd Function", name: 'a')
      book.vba.write("Public Function Gamma()\n  Gamma = 9\nEnd Function", name: 'c', into: 'ThisWorkbook')

      assert_equal 7, xl.ole.Run('Alpha'), 'a standard module answers a bare Run'
      err = assert_raises(WineOLE::RemoteError) { xl.ole.Run('Gamma') }
      assert_match(/0x800A03EC/, err.message, 'ThisWorkbook does not')
    end
  end

  # A codepage file is what Excel's own Export writes and what every .bas
  # from a Windows toolchain is. It used to be refused; it is read now.
  def test_a_codepage_source_file_imports_without_being_re_saved
    with_excel do |xl|
      book = xl['[]']
      Dir.mktmpdir do |dir|
        src = File.join(dir, 'Cp.bas')
        # The module and the procedure must not share a name: VBA resolves
        # the bare name to the module and then reports the macro missing.
        source = +"Attribute VB_Name = \"CpMod\"\r\n" \
                  "Public Function Greeting()\r\n  Greeting = \"αβγ-Дж\"\r\nEnd Function\r\n"
        File.binwrite(src, source.encode(WineOLE::MSOffice::VBA.codepage))
        refute File.binread(src).dup.force_encoding('UTF-8').valid_encoding?,
               'the fixture has to be a file UTF-8 cannot explain, or it proves nothing'
        book.vba.import(src)
        assert_equal "αβγ-Дж", xl.ole.Run('Greeting')
      end
    end
  end

  def test_a_disabled_setting_produces_advice_not_a_com_error
    original = WineOLE::MSOffice::VBA.state
    # Deliberately not a skip. This is the only test that ever exercises the
    # refusal path -- the path most users meet first, and the one this host
    # never takes on its own because the setting is enabled here. If the
    # switch stops working, the suite must say so rather than quietly
    # shipping a branch nobody has run.
    assert WineOLE::MSOffice::VBA.disable!,
      'could not switch AccessVBOM off, so the refusal path cannot be tested'

    begin
      with_excel do |xl|
        err = assert_raises(WineOLE::MSOffice::VBA::AccessDenied) do
          xl['[]'].vba.write('Sub A()', name: 'a')
        end
        assert_match(/wineole-vba enable/, err.message)
      end
    ensure
      original == :enabled ? WineOLE::MSOffice::VBA.enable! : WineOLE::MSOffice::VBA.disable!
    end
  end

  private

  # ExcelIntegrationHelper's own with_excel (used by wineole_integration_test.rb)
  # yields the raw OLE Application -- useful for that file's lifecycle tests,
  # but every test above this line wants the wrapped MSOffice::Excel with a
  # workbook already open. Redefining it here (scoped to this class only,
  # since Ruby methods defined directly on a class win over an included
  # module's) still builds on with_bridge, so the spawn/teardown plumbing
  # itself is not duplicated -- only the "get a wrapped Excel with something
  # open" step that every test below repeats.
  def with_excel
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert do
          xl.ole.Workbooks.Add
          yield xl
        end
      end
    end
  end

  # Every Interior property at once. Color alone cannot tell "no fill" from
  # "painted white" -- both report 16777215.
  def interior_snapshot(range)
    i = range.ole.Interior
    {color_index: i.ColorIndex, pattern: i.Pattern,
     pattern_color_index: i.PatternColorIndex, color: i.Color}
  end
end
