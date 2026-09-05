import contextlib
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.excel_integration import ExcelIntegrationMixin, excel_gone

import wineole
from wineole.msoffice import Book, Color, Excel, Sheet


class MSOfficeIntegrationTest(ExcelIntegrationMixin, unittest.TestCase):
    """Pins the assembled Office wrapper (Address/Paths/Range/Sheet/Book/
    Excel) against real Excel, running under Wine. The unit files test every
    piece against fakes; this is the first place the whole thing is
    exercised end to end."""

    @contextlib.contextmanager
    def wrapped_excel(self):
        """A wrapped Excel with a workbook already open -- what nearly every
        test below wants. Built on `bridge()`, so the spawn/teardown
        plumbing itself is not duplicated."""
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                xl.hide()
                with xl.no_alert():
                    xl.ole.Workbooks().Add()
                    yield xl

    @staticmethod
    def interior_snapshot(rng):
        """Every Interior property at once. Color alone cannot tell "no
        fill" from "painted white" -- both report 16777215."""
        i = rng.ole.Interior()
        return {'color_index': i.ColorIndex(), 'pattern': i.Pattern(),
                'pattern_color_index': i.PatternColorIndex(), 'color': i.Color()}

    # --- lifecycle ---------------------------------------------------------

    def test_run_quits_only_what_it_created(self):
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                self.assertTrue(xl.ole_created,
                                "Excel.run('create') must report the instance as newly created")
                xl.hide()
                with xl.no_alert():
                    xl.ole.Workbooks().Add()

            # Best-effort, not asserted: under Wine the Excel PROCESS
            # sometimes lingers for well over 20s before exiting on its own,
            # with no dialog and nothing this wrapper controls to speed it
            # up. run's finally calling ole_release is what tells the bridge
            # to run CLEANUP_STEPS, and the unit tests pin that call with a
            # fake that records it; asserting on the process here would make
            # this test flake on Wine's own timing.
            if not excel_gone(self.pre_existing_excel_pids, timeout=20):
                print("[note] an Excel process outlived Excel.run('create')'s Quit for over "
                      '20s (known Wine COM release flake)')

            # A second instance, started by hand -- run('connect_or_create')
            # below must attach to this one, not create a third, and must
            # leave it running afterwards.
            xl_raw = client.create('Excel.Application')
            xl_raw.Visible = False
            xl_raw.DisplayAlerts = False
            xl_raw.Workbooks().Add()
            try:
                with Excel.run('connect_or_create', client=client) as xl:
                    self.assertFalse(
                        xl.ole_created,
                        "Excel.run('connect_or_create') must attach to the already-running "
                        'instance, not create one')

                # Still alive: a live Version call is proof, and does not
                # race the way polling a process list would.
                self.assertEqual('11.0', xl_raw.Version(),
                                 "attaching via run('connect_or_create') must not Quit the "
                                 'instance it did not create')
            finally:
                self.quit_bounded(xl_raw)

    def test_leave_open_keeps_excel_running_after_run_and_after_the_connection_closes(self):
        # Asserted after the run block AND after the connection itself is
        # closed: this is only proof if the Excel is still there once the
        # connection that created it is gone. leave_open revokes the
        # bridge's permission to run CLEANUP_STEPS for this record at all.
        #
        # Still inside `bridge()` on purpose, even though the connection is
        # already closed: bridge()'s own finally kills every PID outside the
        # snapshot, so an assertion placed after that block would always
        # find nothing, whether leave_open worked or not. wineole.close() is
        # idempotent, so bridge()'s own close on the way out is a no-op.
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                xl.hide()
                with xl.no_alert():
                    xl.ole.Workbooks().Add()
                xl.leave_open()

            wineole.close()

            # Polled rather than sampled once, in the opposite direction to
            # the usual use: give a cleanup that should NOT happen a full 10
            # seconds to happen, and fail if the process ever disappears.
            self.assertFalse(
                excel_gone(self.pre_existing_excel_pids, timeout=10),
                'leave_open must keep the Excel this run created alive past the block and '
                'the closed connection')
            # Left running deliberately -- bridge()'s own cleanup kills any
            # PID not in the snapshot, so this does not leak.

    # --- the addressing DSL, value shapes, passthrough ----------------------

    def test_addressing_dsl_value_shapes_and_passthrough(self):
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                xl.hide()
                with xl.no_alert():
                    self.assertEqual('11.0', xl.Version(),
                                     'xl.Version() must fall through to COM Application.Version')

                    # New book + new sheet + cell, in one address.
                    xl['[:new]:new!A1'] = 'hello'

                    # ':last!' (bare worksheet, no book part) reaches the
                    # sheet just created -- Worksheets.Add leaves it both
                    # the active sheet and the last one.
                    sheet = xl[':last!']
                    self.assertIsInstance(sheet, Sheet)

                    self.assertEqual([['hello']], sheet['A1'].to_list(),
                                     'Range.to_list must always be 2-D, even for a single cell')
                    self.assertEqual('hello', sheet['A1'].Value(),
                                     "Excel's own Value for a 1x1 range is a bare scalar -- "
                                     'to_list is what normalizes it')

                    # Passthrough to a COM member this wrapper never defines.
                    sheet.ole.PageSetup().Orientation = 2

                    self.assertIsInstance(xl['[]'], Book)

    # --- write vs Excel's own assignment ------------------------------------

    def test_write_lays_a_flat_list_down_the_column_unlike_excels_own_assignment(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':last!']

            sheet['A1:A10'] = list(range(1, 11))
            wrapped = [v for row in sheet['A1:A10'].to_list() for v in row]
            self.assertEqual([float(n) for n in range(1, 11)], wrapped,
                             'Range.write must lay a flat list down a column in order')

            # The identical assignment, done Excel's own way (bypassing
            # write and its shape check entirely via the raw COM Range).
            sheet.ole.Range('B1:B10').Value = list(range(1, 11))
            raw = [v for row in sheet.ole.Range('B1:B10').Value() for v in row]
            self.assertEqual([1.0] * 10, raw,
                             "Excel's own Range.Value= replicates a flat list's first element "
                             'down every cell instead of laying it out -- this is exactly what '
                             'Range.write exists to fix')

            self.assertNotEqual(wrapped, raw,
                                'the whole point of write: the wrapper and the raw COM '
                                'assignment must disagree here')

    def test_write_raises_a_dimension_mismatch_error(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':last!']
            with self.assertRaises(ValueError) as ctx:
                sheet['E1:G3'] = [1, 2]
            self.assertTrue(str(ctx.exception).startswith('range is 3x3;'), str(ctx.exception))
            self.assertIn('only fits a single row or column', str(ctx.exception))

    # --- paths ---------------------------------------------------------------

    def test_save_as_and_local_path_use_linux_paths(self):
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                xl.hide()
                with xl.no_alert():
                    xl['[:new]:new!A1'] = 'save-as check'
                    book = xl['[]']

                    with tempfile.TemporaryDirectory(prefix='wineole-office-integration') as d:
                        out_path = os.path.join(d, 'out.xls')
                        book.save_as(out_path)
                        self.assertTrue(os.path.exists(out_path),
                                        f"save_as({out_path!r}) must create the file at that "
                                        'Linux path')
                        self.assertEqual(d, book.local_path,
                                         'local_path must report the containing folder in '
                                         'Linux form, not a Wine one')
                        book.close(save=False)

    # --- colour --------------------------------------------------------------

    def test_red_is_red_and_not_blue(self):
        # The wrapper's colour vocabulary is RGB; Excel's is BGR. A round
        # trip through the number cannot catch a reversed conversion -- it
        # would agree with itself. Excel's own ColorIndex is what names the
        # colour: 3 is red, 5 is blue.
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            sheet['A1'].format(background='#FF0000')
            self.assertEqual(3, sheet['A1'].ole.Interior().ColorIndex(),
                             'ColorIndex 3 is red; 5 would mean the RGB/BGR conversion is reversed')

            sheet['A2'].format(background='#0000FF')
            self.assertEqual(5, sheet['A2'].ole.Interior().ColorIndex())

    def test_format_reaches_a_whole_range_in_one_call(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            sheet['C1:E3'].format(bold=True, align='center')
            for cell in (sheet['C1'], sheet['E3']):
                self.assertEqual(True, cell.ole.Font().Bold())
                self.assertEqual(-4108, cell.ole.HorizontalAlignment())

    def test_background_false_really_clears_rather_than_painting_white(self):
        # Color cannot express "no fill": a cleared cell and a white-painted
        # cell both report Color 16777215. Asserting on Color alone would
        # accept an implementation that paints white, so this compares every
        # Interior property against a cell that was never touched.
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            pristine = self.interior_snapshot(sheet['Z50'])

            cell = sheet['G1']
            cell.format(background='#FF0000')
            self.assertNotEqual(pristine, self.interior_snapshot(cell),
                                'the fill must actually have happened')

            cell.format(background=False)
            self.assertEqual(pristine, self.interior_snapshot(cell),
                             'a cleared cell must be indistinguishable from one that was '
                             'never filled')

    # --- number_format --------------------------------------------------------

    def test_number_format_general_resets_whatever_the_locale(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            pristine = sheet['Z50'].ole.NumberFormat()

            cell = sheet['G3']
            cell.format(number_format='0.00')
            self.assertNotEqual(pristine, cell.ole.NumberFormat())

            cell.format(number_format='general')
            self.assertEqual(pristine, cell.ole.NumberFormat())

    def test_writing_the_string_general_straight_to_com_still_fails(self):
        # The reason the 'general' special case exists at all. If this ever
        # starts passing, the wrapper is working around something that no
        # longer happens.
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            with self.assertRaises(wineole.RemoteError):
                sheet['G4'].ole.NumberFormat = 'General'

    # --- borders ---------------------------------------------------------------

    def test_borders_reach_the_edges_they_name(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            sheet['I1:K3'].format(border='outline')
            self.assertEqual(1, sheet['I1:K3'].ole.Borders().Item(7).LineStyle())
            self.assertEqual(-4142, sheet['I1:K3'].ole.Borders().Item(11).LineStyle(),
                             "'outline' must leave the inside edges alone")

    def test_border_all_and_false_use_the_bulk_path_without_touching_the_diagonals(self):
        # The bulk path -- an assignment straight to the Borders collection,
        # used for 'all', an explicit list of all six edges, and
        # border=False -- has a danger the 'outline' test never exercises:
        # Excel's Borders collection also holds xlDiagonalDown (index 5) and
        # xlDiagonalUp (6) alongside the six edges this wrapper knows about.
        # If a bulk `Borders.LineStyle =` ever reached those,
        # format(border='all') would silently draw an X through every cell.
        # Measured against a live Excel 11: both diagonals report -4142
        # whether the range has never been formatted or was just formatted
        # with border='all'.
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            rng = sheet['I1:K3']
            borders = rng.ole.Borders()

            for index in range(5, 13):
                self.assertEqual(-4142, borders.Item(index).LineStyle(),
                                 f"edge {index} must start untouched")

            rng.format(border='all')
            self.assertEqual(-4142, borders.Item(5).LineStyle(),
                             "border='all' must leave xlDiagonalDown alone")
            self.assertEqual(-4142, borders.Item(6).LineStyle(),
                             "border='all' must leave xlDiagonalUp alone")
            for index in range(7, 13):
                self.assertEqual(1, borders.Item(index).LineStyle(),
                                 f"border='all' must set edge {index}")

            rng.format(border=False)
            for index in range(5, 13):
                self.assertEqual(-4142, borders.Item(index).LineStyle(),
                                 f"border=False must clear edge {index}")

    def test_excel_2003_snaps_an_off_palette_colour(self):
        # Excel 2003 has a 56-colour palette and silently approximates
        # anything else. Pinned here so the README's warning stays true
        # rather than becoming folklore.
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            sheet['M1'].format(background='#EEEEEE')
            self.assertEqual(Color.parse('#FFFFFF'), int(sheet['M1'].ole.Interior().Color()),
                             'Excel 2003 approximates #EEEEEE to pure white -- not a '
                             'wrapper defect')


if __name__ == '__main__':
    unittest.main()
