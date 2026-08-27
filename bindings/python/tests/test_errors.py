import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.errors import WineOLEError, NotSerializableError, StaleReferenceError, ProtocolError, RemoteError


class ErrorsTest(unittest.TestCase):
    def test_remote_error_message_includes_class_and_message(self):
        err = RemoteError('WIN32OLERuntimeError', 'boom')
        self.assertEqual(str(err), 'WIN32OLERuntimeError: boom')
        self.assertEqual(err.remote_class, 'WIN32OLERuntimeError')

    def test_error_hierarchy(self):
        self.assertTrue(issubclass(NotSerializableError, WineOLEError))
        self.assertTrue(issubclass(StaleReferenceError, WineOLEError))
        self.assertTrue(issubclass(ProtocolError, WineOLEError))
        self.assertTrue(issubclass(RemoteError, WineOLEError))


if __name__ == '__main__':
    unittest.main()
