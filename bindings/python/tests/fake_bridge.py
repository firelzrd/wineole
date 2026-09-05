"""An in-process stand-in for wineole-bridge, on a real loopback socket.

A real socket rather than a fake object, because what the unit tests below it
exercise IS the reader thread and the framing: a stand-in for the socket would
test neither. It records what it was asked, answers either automatically (a
`handler`) or by hand (`take_request` + `reply`, for tests about two requests
in flight at once or answers arriving out of order), and can push unsolicited
event frames and plain garbage lines the way the bridge does.
"""

import json
import queue
import socket
import threading
import traceback


class Refusal:
    """What an auto-answer handler returns to make the bridge refuse a request.

    The wire has two shapes of answer, `result` and `error`, and a handler
    returning a plain dict could not say which one it meant -- a result IS a
    dict for nearly every method. This names the error shape instead.
    """

    __slots__ = ('class_name', 'message')

    def __init__(self, class_name='WIN32OLERuntimeError', message='refused'):
        self.class_name = class_name
        self.message = message


class FakeBridge:
    """Use as a context manager; `bridge.sock` is the client end.

    With `handler=None` nothing is answered automatically: requests land in an
    inbox for `take_request`, and the test answers with `reply`. With a
    handler, every request is answered with `handler(method, params)` -- a
    `Refusal` becomes an error frame, anything else becomes a result frame.

    If `handler` raises, the exception is appended to `handler_errors` (a
    public list), printed to stderr with a traceback, and the request is
    answered with an error frame (class `FakeBridgeHandlerError`) instead of
    being left to hang -- a broken handler must fail its client call
    visibly, not wedge the serve loop for every request after it.
    """

    def __init__(self, handler=None):
        self._handler = handler
        self.requests = []
        self.handler_errors = []
        self._inbox = queue.Queue()
        self._listener = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self._listener.bind(('127.0.0.1', 0))
        self._listener.listen(1)
        self._conn = None
        self._accepted = threading.Event()
        self._write_lock = threading.Lock()
        self._thread = None
        self.sock = None

    def __enter__(self):
        self._thread = threading.Thread(target=self._serve, name='fake-bridge', daemon=True)
        self._thread.start()
        self.sock = socket.create_connection(self._listener.getsockname())
        if not self._accepted.wait(5):
            raise AssertionError('the fake bridge never accepted the connection')
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        try:
            self.sock.close()
        except OSError:
            pass
        self._thread.join(5)
        return False

    def _serve(self):
        try:
            conn, _ = self._listener.accept()
        except OSError:
            return
        self._conn = conn
        self._accepted.set()
        reader = conn.makefile('rb')
        try:
            while True:
                line = reader.readline()
                if not line:
                    break
                request = json.loads(line)
                # list.append is atomic under the GIL, so a test thread may
                # read `requests` while this one writes it.
                self.requests.append((request.get('method'), request.get('params')))
                self._inbox.put(request)
                if self._handler is None:
                    continue
                try:
                    answer = self._handler(request.get('method'), request.get('params'))
                except Exception as exc:
                    # A handler that raises must not wedge the serve loop:
                    # every client call after it would otherwise hang
                    # waiting for a reply that never comes.
                    self.handler_errors.append(exc)
                    traceback.print_exc()
                    self.push({'id': request.get('id'), 'error': {
                        'class': 'FakeBridgeHandlerError',
                        'message': repr(exc),
                    }})
                    continue
                if isinstance(answer, Refusal):
                    self.reply_error(request, answer.class_name, answer.message)
                else:
                    self.reply(request, answer)
        except (OSError, ValueError):
            # The client closed the socket under it, or sent a line that is
            # not a request. Either way this bridge is done.
            pass
        finally:
            try:
                reader.close()
            except OSError:
                pass

    def take_request(self, timeout=5):
        """The next request as a dict, or None if none arrived in time.

        Bounded and None-on-miss rather than blocking: the bug most of these
        tests are hunting is a request that never reaches the wire, and an
        unbounded wait would hang the suite instead of failing it.
        """
        try:
            return self._inbox.get(timeout=timeout)
        except queue.Empty:
            return None

    def reply(self, request, result):
        self.push({'id': request['id'], 'result': result})

    def reply_error(self, request, class_name, message):
        self.push({'id': request['id'], 'error': {'class': class_name, 'message': message}})

    def push(self, frame):
        """Write one frame to the client, exactly as the bridge would."""
        self.push_line(json.dumps(frame))

    def push_line(self, raw):
        """Write one raw line -- garbage included -- to the client."""
        with self._write_lock:
            conn = self._conn
            if conn is None:
                raise AssertionError('the fake bridge has no connection to push on')
            try:
                conn.sendall((raw + '\n').encode('utf-8'))
            except OSError:
                pass

    def close_write(self):
        """Half-close: the client sees EOF while its own socket stays writable."""
        if self._conn is not None:
            try:
                self._conn.shutdown(socket.SHUT_WR)
            except OSError:
                pass

    def close(self):
        """End the stream."""
        if self._conn is not None:
            try:
                # shutdown before close: the serve thread's makefile('rb')
                # keeps the fd alive after close() alone, so the peer would
                # never see EOF and would hang waiting for one.
                self._conn.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
            try:
                self._conn.close()
            except OSError:
                pass
        try:
            self._listener.close()
        except OSError:
            pass

    def count(self, method):
        return sum(1 for name, _ in self.requests if name == method)
