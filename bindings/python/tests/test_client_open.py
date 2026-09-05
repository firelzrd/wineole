import json
import os
import socket
import sys
import tempfile
import threading
import time
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.client import Client
from wineole.errors import WineOLEError


class ClientOpenTest(unittest.TestCase):
    def test_reuses_already_listening_server_without_spawning(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]
        spawn_called = []

        def spawner(p):
            spawn_called.append(p)

        # open always pings to handshake, even on the
        # already-listening fast path (mirrors client.rb's
        # test_pings_after_connecting_to_an_existing_bridge) -- so this fake
        # server must actually accept and answer, or the handshake's
        # readline() blocks forever.
        def run_server():
            conn, _ = server.accept()
            line = conn.makefile().readline()
            req = json.loads(line)
            conn.sendall((json.dumps({'id': req['id'], 'result': {'pong': True}}) + '\n').encode())
            conn.close()

        thread = threading.Thread(target=run_server)
        thread.start()

        with tempfile.TemporaryDirectory() as tmpdir:
            client = Client.open(
                port=port, spawner=spawner, lockfile=os.path.join(tmpdir, 'lock')
            )
            self.assertEqual(spawn_called, [])
            client.close()
        thread.join()
        server.close()

    def test_spawns_when_nothing_listening(self):
        # Below 32768, outside Linux's default ephemeral range (32768-60999),
        # so this bind cannot collide with a kernel-assigned outbound port.
        port = 21000 + (os.getpid() % 1000)
        spawned_port = []

        def spawner(p):
            spawned_port.append(p)

            # As above: the post-spawn path also handshakes, so this must
            # actually answer the ping (mirrors client.rb's
            # test_pings_after_spawning_a_new_bridge), not merely accept().
            def start_server():
                server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                server.bind(('127.0.0.1', p))
                server.listen(1)
                conn, _ = server.accept()
                line = conn.makefile().readline()
                req = json.loads(line)
                conn.sendall((json.dumps({'id': req['id'], 'result': {'pong': True}}) + '\n').encode())
                conn.close()
                server.close()

            threading.Thread(target=start_server, daemon=True).start()

        with tempfile.TemporaryDirectory() as tmpdir:
            client = Client.open(
                port=port, spawner=spawner, lockfile=os.path.join(tmpdir, 'lock'), timeout=5
            )
            self.assertEqual(spawned_port, [port])
            client.close()

    def test_raises_if_spawned_server_never_comes_up(self):
        # Below 32768, outside Linux's default ephemeral range -- see above.
        port = 22000 + (os.getpid() % 1000)

        def spawner(p):
            pass  # nobody ever listens

        with tempfile.TemporaryDirectory() as tmpdir:
            with self.assertRaises(WineOLEError):
                Client.open(
                    port=port, spawner=spawner, lockfile=os.path.join(tmpdir, 'lock'), timeout=1
                )

    def test_bridge_path_for_arch_raises_for_unsupported_architecture(self):
        with self.assertRaises(WineOLEError) as ctx:
            Client.bridge_path_for_arch('sparc64')
        self.assertIn('sparc64', str(ctx.exception))


if __name__ == '__main__':
    unittest.main()
