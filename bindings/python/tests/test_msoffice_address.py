import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.msoffice.address import Address


def parse(text, version=11.0):
    return Address.parse(text, version)


def column_names(count):
    """A..Z, AA..ZZ, AAA..ZZZ in Excel's own order, truncated to `count`."""
    letters = [chr(c) for c in range(ord('A'), ord('Z') + 1)]
    out = list(letters)
    for a in letters:
        for b in letters:
            out.append(a + b)
    for a in letters:
        for b in letters:
            for c in letters:
                out.append(a + b + c)
    return out[:count]


class MSOfficeAddressTest(unittest.TestCase):
    def test_parses_a_bare_range(self):
        a = parse('A1:B2')
        self.assertIsNone(a.workbook)
        self.assertIsNone(a.worksheet)
        self.assertEqual('A1:B2', a.range)
        self.assertTrue(a.has_range)

    def test_parses_a_worksheet_and_range(self):
        a = parse('Sheet1!A1')
        self.assertEqual('Sheet1', a.worksheet)
        self.assertEqual('A1', a.range)

    def test_parses_a_workbook_worksheet_and_range(self):
        a = parse('[Book1]Sheet1!A1:B2')
        self.assertEqual('Book1', a.workbook)
        self.assertEqual('Sheet1', a.worksheet)
        self.assertEqual('A1:B2', a.range)

    def test_parses_the_new_workbook_marker(self):
        a = parse('[:new]')
        self.assertEqual(':new', a.workbook)
        self.assertIsNone(a.range)
        self.assertFalse(a.has_range)

    def test_parses_the_sheet_markers(self):
        self.assertEqual(':new', parse(':new!').worksheet)
        self.assertEqual(':first', parse(':first!A1').worksheet)
        self.assertEqual(':last', parse(':last!A1').worksheet)

    def test_an_empty_workbook_or_worksheet_means_the_active_one(self):
        self.assertEqual('', parse('[]').workbook)
        self.assertEqual('', parse('!A1').worksheet)

    def test_a_sheet_without_a_range_is_not_assignable(self):
        a = parse('Sheet1!')
        self.assertEqual('Sheet1', a.worksheet)
        self.assertIsNone(a.range)
        self.assertFalse(a.has_range,
                         'a sheet address carries no range, so it must not accept an assignment')

    def test_rejects_a_row_beyond_the_excel_11_limit(self):
        self.assertIsNone(parse('A65537'), 'row 65537 does not exist in Excel 11')
        self.assertIsNotNone(parse('A65536'))

    def test_rejects_a_column_beyond_the_excel_11_limit(self):
        self.assertIsNotNone(parse('IV65536'), 'IV65536 is the last cell of the Excel 11 grid')
        self.assertIsNone(parse('IW1'), 'IW is past the last column of the Excel 11 grid')

    def test_excel_12_allows_the_larger_grid(self):
        self.assertIsNone(parse('A65537', 11.0))
        self.assertIsNotNone(parse('A65537', 12.0))
        self.assertEqual('XFD1048576', parse('XFD1048576', 12.0).range)
        self.assertIsNone(parse('XFE1', 12.0), 'XFE is past the last column')

    # Checking a handful of boundaries would not have caught what was
    # actually wrong with the Ruby patterns: row 65530 alone was missing
    # from the Excel 11 grid, and 11,111 scattered rows from the Excel 12
    # one. Both grids are small enough to check in full.
    def test_every_row_in_the_grid_parses(self):
        for version, last_row in ((11.0, 65536), (12.0, 1048576)):
            missing = [n for n in range(1, last_row + 1) if parse(f"A{n}", version) is None]
            self.assertEqual([], missing[:10],
                             f"{len(missing)} row(s) of the Excel {int(version)} grid do not parse")
            past_the_end = [n for n in range(last_row + 1, last_row + 21)
                            if parse(f"A{n}", version) is not None]
            self.assertEqual([], past_the_end,
                             f"rows past the end of the Excel {int(version)} grid must not parse")

    def test_every_column_in_the_grid_parses(self):
        for version, last, count in ((11.0, 'IV', 256), (12.0, 'XFD', 16384)):
            cols = column_names(count)
            self.assertEqual(last, cols[-1], 'the column list must end at the grid limit')
            missing = [c for c in cols if parse(f"{c}1", version) is None]
            self.assertEqual([], missing[:10],
                             f"{len(missing)} column(s) of the Excel {int(version)} grid do not parse")

    def test_returns_none_for_something_that_is_not_an_address(self):
        self.assertIsNone(parse('this is not an address'),
                          'the caller treats None as "not an address" and falls back to a raw sheet name')


if __name__ == '__main__':
    unittest.main()
