import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import FakeComWorksheet

from wineole.msoffice.range import Range
from wineole.msoffice.sheet import Sheet


def make_sheet(version='11.0'):
    return Sheet(FakeComWorksheet('Sheet1'), version)


class MSOfficeSheetTest(unittest.TestCase):
    # --- [] with two integers -> Cells --------------------------------

    def test_subscript_with_two_integers_wraps_cells(self):
        s = make_sheet()
        r = s[2, 3]
        self.assertIsInstance(r, Range)
        self.assertEqual([[2, 3]], s.ole.cells_calls)

    # --- [] with a string address -> Range -----------------------------

    def test_subscript_with_a_string_address_wraps_range(self):
        s = make_sheet()
        r = s['A1:B2']
        self.assertIsInstance(r, Range)
        self.assertEqual(['A1:B2'], s.ole.range_calls)

    # --- version is plumbed through to Address.parse -------------------

    def test_version_controls_which_grid_is_accepted(self):
        # XFD1 is beyond Excel 11's IV/65536 grid, but within Excel 12's.
        with self.assertRaises(ValueError) as ctx:
            make_sheet('11.0')['XFD1']
        self.assertIn('range', str(ctx.exception))

        self.assertIsInstance(make_sheet('12.0')['XFD1'], Range)

    # --- []= delegates to Range.write, never fill -----------------------

    def test_subscript_assign_writes_through_range_write(self):
        s = make_sheet()
        s['A1:B1'] = 7
        self.assertEqual(7, s.ole.ranges[-1].written)

    def test_subscript_assign_calls_write_not_fill(self):
        s = make_sheet()
        # write() raises on a flat list that does not exactly fit the range;
        # fill() would happily replicate/pad it instead and never raise. The
        # fake range always reports itself as 1x1, so any multi-element list
        # proves it is write(), not fill(), backing []=.
        with self.assertRaises(ValueError) as ctx:
            s['A1:B1'] = [1, 2, 3]
        self.assertIn('1x1', str(ctx.exception))
        self.assertIn('3 elements', str(ctx.exception))

    def test_subscript_assign_with_two_integers(self):
        s = make_sheet()
        s[1, 1] = 5
        self.assertEqual([[1, 1]], s.ole.cells_calls)
        self.assertEqual(5, s.ole.ranges[-1].written)

    # --- addresses without a range raise, both for read and for write ---

    def test_assigning_to_an_address_without_a_range_raises(self):
        s = make_sheet()
        with self.assertRaises(ValueError) as ctx:
            s[''] = 0
        self.assertIn('range', str(ctx.exception))

    def test_reading_an_address_without_a_range_also_raises(self):
        s = make_sheet()
        with self.assertRaises(ValueError) as ctx:
            s['']
        self.assertIn('range', str(ctx.exception))

    # --- an address naming another sheet is refused --------------------

    def test_an_address_naming_another_worksheet_raises(self):
        s = make_sheet()
        with self.assertRaises(ValueError) as ctx:
            s['Sheet2!A1']
        self.assertIn('sheet', str(ctx.exception).lower())

    def test_an_address_naming_a_workbook_raises(self):
        s = make_sheet()
        with self.assertRaises(ValueError) as ctx:
            s['[Book2]Sheet1!A1']
        self.assertIn('sheet', str(ctx.exception).lower())

    # --- an index of the wrong shape ------------------------------------

    def test_an_unsupported_index_raises_a_type_error(self):
        s = make_sheet()
        for bad in [1.5, (1, 2, 3), None, (1, 'A')]:
            with self.assertRaises(TypeError, msg=f"{bad!r} must be refused"):
                s[bad]

    # --- passthrough ------------------------------------------------------

    def test_unknown_methods_go_to_com(self):
        self.assertEqual('Sheet1', make_sheet().Name())

    def test_ole_exposes_the_underlying_proxy(self):
        s = make_sheet()
        self.assertIsInstance(s.ole, FakeComWorksheet)


if __name__ == '__main__':
    unittest.main()
