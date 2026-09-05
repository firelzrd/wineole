import os
import re
import shutil
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import FakeClient, FakeComWorkbook, FakeComWorksheet

from wineole.msoffice.book import Book
from wineole.msoffice.sheet import Sheet

WINDOWS_SHAPED = re.compile(r'^(?:[A-Za-z]:[\\/]|\\\\)')


def has_winepath():
    return shutil.which('winepath') is not None


def make_book(ole=None, loopback=True, convert_paths=True, version='11.0'):
    if ole is None:
        ole = FakeComWorkbook(worksheets=[FakeComWorksheet('Sheet1'),
                                         FakeComWorksheet('Sheet2')])
    return Book(ole, FakeClient(loopback=loopback), version, convert_paths=convert_paths)


class MSOfficeBookTest(unittest.TestCase):
    # --- save_as: path conversion is gated correctly ----------------------

    def test_a_remote_clients_save_as_passes_the_path_unconverted(self):
        ole = FakeComWorkbook()
        make_book(ole=ole, loopback=False).save_as('/home/user/out.xls')
        self.assertEqual(['/home/user/out.xls'], ole.save_as_calls)

    def test_convert_paths_false_on_a_loopback_client_passes_the_path_unconverted(self):
        ole = FakeComWorkbook()
        make_book(ole=ole, loopback=True, convert_paths=False).save_as('/home/user/out.xls')
        self.assertEqual(['/home/user/out.xls'], ole.save_as_calls)

    def test_a_loopback_clients_save_as_hands_com_a_windows_shaped_path(self):
        if not has_winepath():
            self.skipTest('winepath is not on PATH')

        ole = FakeComWorkbook()
        make_book(ole=ole, loopback=True, convert_paths=True).save_as('/home/user/out.xls')
        self.assertEqual(1, len(ole.save_as_calls))
        self.assertRegex(ole.save_as_calls[0], WINDOWS_SHAPED,
                         'a loopback client should get a Windows-shaped path back from winepath')

    # --- convert_paths cannot be talked into converting -------------------

    def test_a_remote_client_cannot_be_talked_into_converting(self):
        ole = FakeComWorkbook()
        make_book(ole=ole, loopback=False, convert_paths=True).save_as('/home/user/out.xls')
        self.assertEqual(['/home/user/out.xls'], ole.save_as_calls,
                         'a caller passing convert_paths=True must not override a remote '
                         "bridge's own answer about whether converting means anything")

    # --- local_path -------------------------------------------------------

    def test_local_path_is_left_alone_when_not_converting(self):
        ole = FakeComWorkbook(path='Z:\\home\\user')
        self.assertEqual('Z:\\home\\user', make_book(ole=ole, loopback=False).local_path)

    def test_local_path_of_an_unsaved_book_is_empty_without_shelling_out(self):
        ole = FakeComWorkbook(path='')
        self.assertEqual('', make_book(ole=ole, loopback=True, convert_paths=True).local_path)

    # --- local_file -------------------------------------------------------
    #
    # Gated the same way local_path is: an ungated
    # Paths.to_local(book.FullName()) would, over a remote bridge, silently
    # run a local winepath over a path that names some other machine's
    # filesystem.

    def test_a_remote_clients_local_file_is_left_unconverted(self):
        ole = FakeComWorkbook(full_name='Z:\\tmp\\out.xls')
        self.assertEqual('Z:\\tmp\\out.xls', make_book(ole=ole, loopback=False).local_file)

    def test_a_loopback_clients_local_file_is_a_linux_path(self):
        if not has_winepath():
            self.skipTest('winepath is not on PATH')

        ole = FakeComWorkbook(full_name='Z:\\tmp\\out.xls')
        self.assertNotRegex(make_book(ole=ole, loopback=True, convert_paths=True).local_file,
                            WINDOWS_SHAPED,
                            'a loopback client should get a Linux path back from winepath')

    def test_local_file_of_an_unsaved_book_does_not_shell_out(self):
        ole = FakeComWorkbook(full_name='Book1')
        self.assertEqual('Book1', make_book(ole=ole, loopback=True, convert_paths=True).local_file)

    # --- sheet / sheets go through Worksheets, not Sheets ------------------

    def test_sheet_wraps_a_worksheet_by_name_or_index(self):
        ole = FakeComWorkbook(worksheets=[FakeComWorksheet('Sheet1'), FakeComWorksheet('Sheet2')])
        b = make_book(ole=ole)
        s = b.sheet('Sheet1')
        self.assertIsInstance(s, Sheet)
        self.assertEqual('Sheet1', s.ole.name)
        self.assertEqual(['Sheet1'], ole.worksheets.item_calls,
                         'Worksheets, not Sheets: Sheets also holds chart sheets, which '
                         'this wrapper does not model')

    def test_sheets_yields_one_sheet_per_worksheet(self):
        b = make_book()
        got = list(b.sheets())
        self.assertEqual(2, len(got))
        self.assertTrue(all(isinstance(s, Sheet) for s in got))

    def test_sheets_is_lazy(self):
        # A generator, not a list: nothing is fetched until it is iterated,
        # which is what Ruby's each_sheet enumerator gives too.
        ole = FakeComWorkbook(worksheets=[FakeComWorksheet('Sheet1'), FakeComWorksheet('Sheet2')])
        b = make_book(ole=ole)
        gen = b.sheets()
        self.assertEqual([], ole.worksheets.item_calls)
        next(gen)
        self.assertEqual([1], ole.worksheets.item_calls)

    # --- close is a deliberate shadow, not the raw COM member -------------

    def test_close_defaults_to_not_saving(self):
        ole = FakeComWorkbook()
        make_book(ole=ole).close()
        self.assertEqual([False], ole.close_calls)

    def test_close_can_be_told_to_save(self):
        ole = FakeComWorkbook()
        make_book(ole=ole).close(save=True)
        self.assertEqual([True], ole.close_calls)

    # --- the raw COM member stays reachable in PascalCase ------------------

    def test_the_raw_com_close_stays_reachable_in_pascal_case(self):
        ole = FakeComWorkbook()
        make_book(ole=ole).Close(True)
        self.assertEqual([True], ole.close_calls)


if __name__ == '__main__':
    unittest.main()
