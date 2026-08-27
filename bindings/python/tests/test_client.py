import gc
import json
import os
import socket
import sys
import threading
import unittest
import weakref

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.client import Client
from wineole.errors import RemoteError, ProtocolError
from wineole.proxy import Proxy


class ClientTest(unittest.TestCase):
    def test_call_returns_result_on_success(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]

        def run_server():
            conn, _ = server.accept()
            line = conn.makefile().readline()
            req = json.loads(line)
            conn.sendall((json.dumps({'id': req['id'], 'result': {'pong': True}}) + '\n').encode())
            conn.close()

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        try:
            result = client.call('ping', {})
        finally:
            client.close()

        self.assertEqual(result, {'pong': True})
        thread.join()
        server.close()

    def test_call_raises_remote_error_on_error_response(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]

        def run_server():
            conn, _ = server.accept()
            line = conn.makefile().readline()
            req = json.loads(line)
            error = {'class': 'WIN32OLERuntimeError', 'message': 'boom'}
            conn.sendall((json.dumps({'id': req['id'], 'error': error}) + '\n').encode())
            conn.close()

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        try:
            with self.assertRaises(RemoteError) as ctx:
                client.call('invoke', {})
            self.assertEqual(str(ctx.exception), 'WIN32OLERuntimeError: boom')
        finally:
            client.close()

        thread.join()
        server.close()

    def test_each_client_gets_a_distinct_never_reused_generation(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(2)
        port = server.getsockname()[1]

        def accept_and_close(n):
            for _ in range(n):
                conn, _ = server.accept()
                conn.close()

        thread = threading.Thread(target=accept_and_close, args=(2,))
        thread.start()

        client_a = Client(socket.create_connection(('127.0.0.1', port)))
        client_b = Client(socket.create_connection(('127.0.0.1', port)))

        try:
            # Consecutive, not merely distinct. Two *simultaneously live*
            # objects always have distinct id()s, so a bare assertNotEqual
            # would still pass if `generation` were reverted to `id(self)` —
            # the Ruby-style pattern this client deliberately avoids, because
            # CPython reuses an id() once its object is collected, while
            # Ruby's object_id never is. Asserting that client_b lands on
            # exactly the next value pins the monotonic-counter mechanism
            # itself, which id() cannot satisfy. No Client is constructed
            # between these two lines, so the shared module-level counter
            # cannot be bumped by anything else in between.
            self.assertEqual(client_b.generation, client_a.generation + 1)
        finally:
            client_a.close()
            client_b.close()

        thread.join()
        server.close()

    def test_concurrent_calls_from_multiple_threads_do_not_interleave(self):
        # Regression test for the threading.Lock added to call(): without it,
        # concurrent threads sharing one Client can interleave their
        # sendall()/readline() pairs, corrupting request framing and causing
        # spurious "id mismatch" ProtocolErrors or a thread reading another
        # thread's response. The server below sleeps between reading a
        # request and replying, widening the interleaving window so a
        # missing lock fails this test reliably rather than by chance.
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]

        def run_server():
            conn, _ = server.accept()
            reader = conn.makefile('r')
            while True:
                line = reader.readline()
                if not line:
                    break
                req = json.loads(line)
                threading.Event().wait(0.002)
                reply = {'id': req['id'], 'result': {'echo': req['params']['n']}}
                conn.sendall((json.dumps(reply) + '\n').encode())
            conn.close()

        server_thread = threading.Thread(target=run_server)
        server_thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)

        errors = []
        CALLS_PER_WORKER = 30

        def worker(worker_id):
            try:
                for i in range(CALLS_PER_WORKER):
                    n = worker_id * 1000 + i
                    result = client.call('echo', {'n': n})
                    if result != {'echo': n}:
                        errors.append(f"worker {worker_id} call {i}: got {result!r}, expected echo of {n}")
            except Exception as exc:
                errors.append(f"worker {worker_id}: {exc!r}")

        workers = [threading.Thread(target=worker, args=(w,)) for w in range(8)]
        for w in workers:
            w.start()
        for w in workers:
            w.join(timeout=10)
            self.assertFalse(w.is_alive(), "worker thread did not finish — likely deadlocked")

        client.close()
        server_thread.join(timeout=5)
        server.close()

        self.assertEqual(errors, [])

    def test_create_delegates_to_proxy_create(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]

        def run_server():
            conn, _ = server.accept()
            line = conn.makefile().readline()
            req = json.loads(line)
            conn.sendall((json.dumps({'id': req['id'], 'result': {'$ole_ref': 7}}) + '\n').encode())
            conn.close()

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        try:
            proxy = client.create('Excel.Application')
            self.assertIsInstance(proxy, Proxy)
            self.assertEqual(proxy.ole_handle, 7)
        finally:
            client.close()

        thread.join()
        server.close()

    def test_connect_delegates_to_proxy_connect(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]

        def run_server():
            conn, _ = server.accept()
            line = conn.makefile().readline()
            req = json.loads(line)
            conn.sendall((json.dumps({'id': req['id'], 'result': {'$ole_ref': 9}}) + '\n').encode())
            conn.close()

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        try:
            proxy = client.connect('Excel.Application')
            self.assertIsInstance(proxy, Proxy)
            self.assertEqual(proxy.ole_handle, 9)
        finally:
            client.close()

        thread.join()
        server.close()

    def test_connect_or_create_delegates_to_proxy_connect_or_create(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]

        def run_server():
            conn, _ = server.accept()
            line = conn.makefile().readline()
            req = json.loads(line)
            conn.sendall((json.dumps({'id': req['id'], 'result': {'$ole_ref': 11, 'created': False}}) + '\n').encode())
            conn.close()

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        try:
            proxy = client.connect_or_create('Excel.Application')
            self.assertIsInstance(proxy, Proxy)
            self.assertEqual(proxy.ole_handle, 11)
            self.assertFalse(proxy.ole_created)
        finally:
            client.close()

        thread.join()
        server.close()

    def test_close_does_not_raise_when_called_twice(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]
        accepted = []

        def run_server():
            conn, _ = server.accept()
            accepted.append(conn)

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        thread.join()

        client.close()
        client.close()

        accepted[0].close()
        server.close()

    def test_del_closes_the_socket(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server.bind(('127.0.0.1', 0))
        server.listen(1)
        port = server.getsockname()[1]
        accepted = []

        def run_server():
            conn, _ = server.accept()
            accepted.append(conn)

        thread = threading.Thread(target=run_server)
        thread.start()

        sock = socket.create_connection(('127.0.0.1', port))
        client = Client(sock)
        thread.join()

        weak = weakref.ref(client)
        client = None
        gc.collect()

        self.assertIsNone(weak(), 'the Client must actually be collected (no reference cycle from __del__)')

        conn = accepted[0]
        conn.settimeout(2)
        self.assertEqual(conn.recv(1), b'', 'the peer must see EOF once __del__ closed the underlying socket')
        conn.close()
        server.close()


if __name__ == '__main__':
    unittest.main()
