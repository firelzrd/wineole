import os
import shutil
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.msoffice.paths import Paths


class FakeClientForPaths:
    """Only what Paths.convertible asks of a client: the `loopback` property
    the real Client exposes."""

    def __init__(self, loopback):
        self.loopback = loopback


def has_winepath():
    return shutil.which('winepath') is not None


class MSOfficePathsTest(unittest.TestCase):
    # --- convertible ----------------------------------------------------

    def test_paths_are_not_converted_when_the_client_runs_on_windows(self):
        self.assertFalse(
            Paths.convertible(FakeClientForPaths(True), windows=True),
            'a Windows client already has Windows paths, and there is no winepath')

    def test_paths_are_not_converted_for_a_remote_bridge(self):
        self.assertFalse(
            Paths.convertible(FakeClientForPaths(False), windows=False),
            "the client's own path means nothing on another machine")

    def test_paths_are_converted_for_a_local_bridge(self):
        self.assertTrue(Paths.convertible(FakeClientForPaths(True), windows=False))

    # --- to_wine: already-Windows paths are left alone -------------------

    def test_an_already_windows_path_is_left_alone(self):
        self.assertEqual('Z:\\home\\user\\out.xls', Paths.to_wine('Z:\\home\\user\\out.xls'))
        self.assertEqual('C:\\Temp\\x.xls', Paths.to_wine('C:\\Temp\\x.xls'))
        self.assertEqual('\\\\server\\share\\x.xls', Paths.to_wine('\\\\server\\share\\x.xls'))

    # --- to_local: already-local paths are left alone --------------------

    def test_a_non_windows_shaped_path_is_left_alone_by_to_local(self):
        self.assertEqual('/home/user/out.xls', Paths.to_local('/home/user/out.xls'))

    def test_an_empty_path_is_returned_unchanged(self):
        # An unsaved Workbook's Path is "" -- neither direction must shell
        # out for it.
        self.assertEqual('', Paths.to_local(''))
        self.assertEqual('', Paths.to_wine(''))

    # --- round trip via the real winepath, skipped if it is not installed -

    def test_a_linux_path_round_trips_through_winepath(self):
        if not has_winepath():
            self.skipTest('winepath is not on PATH')

        # A drive root (e.g. the Z: mapping of a whole prefix) round-trips
        # with a trailing slash added by winepath itself, so exercise a
        # subdirectory instead -- the round trip this method promises.
        path = os.path.dirname(os.path.abspath(__file__))
        wine_path = Paths.to_wine(path)
        self.assertEqual(path, Paths.to_local(wine_path))


if __name__ == '__main__':
    unittest.main()
