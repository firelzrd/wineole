import json
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from fake_bridge import FakeBridge


# FakeBridge itself has no behaviour of its own worth testing beyond this:
# Tasks 2-6 write handlers, and a handler that raises must fail the request
# visibly instead of wedging the serve loop for every call after it.
class FakeBridgeHandlerErrorTest(unittest.TestCase):
    def test_a_raising_handler_answers_with_an_error_frame_instead_of_hanging(self):
        def handler(method, params):
            raise ValueError('boom')

        with FakeBridge(handler=handler) as bridge:
            bridge.sock.sendall(json.dumps({'id': 1, 'method': 'anything', 'params': {}}).encode('utf-8') + b'\n')
            reader = bridge.sock.makefile('rb')
            line = reader.readline()
            self.assertTrue(line, 'the client must get an answer, not a hang')
            frame = json.loads(line)

            self.assertEqual(frame['id'], 1)
            self.assertEqual(frame['error']['class'], 'FakeBridgeHandlerError')
            self.assertIn('boom', frame['error']['message'])

            self.assertEqual(len(bridge.handler_errors), 1)
            self.assertIsInstance(bridge.handler_errors[0], ValueError)

            # Close the client end from this side before the `with` block's
            # __exit__ runs: fake_bridge's close() does not shut the socket
            # down before closing it, so the serve thread's blocking read
            # would otherwise sit until __exit__'s 5s join gives up.
            reader.close()
            bridge.sock.close()


if __name__ == '__main__':
    unittest.main()
