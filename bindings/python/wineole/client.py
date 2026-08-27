import itertools
import json
import os
import platform
import socket
import subprocess
import tempfile
import threading
import time

# fcntl (POSIX) and msvcrt (Windows) are both stdlib but each is only
# importable on its own platform -- there's no third module both understand,
# so the client has to actually branch on OS instead of picking one.
if os.name == 'nt':
    import msvcrt
else:
    import fcntl

from .errors import WineOLEError, ProtocolError, RemoteError
from .proxy import Proxy

_client_generation = itertools.count(1)

# Mirrors wineole/client.rb's ARCH_TRIPLES (Ruby, Task 5) — kept in sync by
# hand, since there is no shared config file between the two clients.
ARCH_TRIPLES = {
    'x86_64': 'x86_64-pc-windows-gnu',
    'amd64': 'x86_64-pc-windows-gnu',
    'i386': 'i686-pc-windows-gnu',
    'i686': 'i686-pc-windows-gnu',
}


class Client:
    def __init__(self, sock):
        self._sock = sock
        self._reader = sock.makefile('r')
        self._next_id = 0
        self.generation = next(_client_generation)
        # Mirrors wineole/client.rb's `@mutex = Mutex.new`: one Client may be
        # shared across threads, and `call` is a read-modify-write on
        # `_next_id` followed by a send/receive pair that must not interleave
        # with another thread's — otherwise a caller gets a spurious
        # `id mismatch` ProtocolError, or silently receives another thread's
        # response.
        self._lock = threading.Lock()

    def __del__(self):
        # Safety net for a forgotten close() -- best-effort, not a
        # replacement for it (GC timing is never guaranteed). Exceptions
        # must not escape __del__: the interpreter would just print them to
        # stderr and continue, so there's nothing to gain by letting one
        # propagate, and it would be noisy for no benefit.
        try:
            self.close()
        except Exception:
            pass

    def call(self, method, params=None):
        with self._lock:
            params = params if params is not None else {}
            self._next_id += 1
            request_id = self._next_id
            self._sock.sendall((json.dumps({'id': request_id, 'method': method, 'params': params}) + '\n').encode())
            line = self._reader.readline()
            if not line:
                raise ProtocolError('connection closed')
            response = json.loads(line)
            if response.get('id') != request_id:
                raise ProtocolError(f"id mismatch: expected {request_id}, got {response.get('id')}")
            if response.get('error'):
                raise RemoteError(response['error']['class'], response['error']['message'])
            return response.get('result')

    def close(self):
        # Closing only the socket is not enough: makefile() holds an io-ref on
        # it, so socket.close() defers the real fd close until the reader is
        # finalized. Close the reader first so the peer actually sees EOF —
        # which is what triggers the bridge's session/handle reclamation
        # (handle_connection -> release_all). Without this, a live Proxy
        # (which holds a reference to its Client) keeps the connection — and
        # therefore the COM session and its automated application — alive
        # indefinitely.
        try:
            self._reader.close()
        finally:
            self._sock.close()

    def create(self, class_name):
        return Proxy.create(class_name, self)

    def connect(self, class_name):
        return Proxy.connect(class_name, self)

    def connect_or_create(self, class_name):
        return Proxy.connect_or_create(class_name, self)

    @staticmethod
    def bridge_path_for_arch(machine):
        triple = ARCH_TRIPLES.get(machine)
        if triple is None:
            raise WineOLEError(
                f"no prebuilt wineole-bridge binary for host architecture {machine!r} "
                f"(available: {sorted(set(ARCH_TRIPLES.values()))})"
            )
        return os.path.abspath(
            os.path.join(os.path.dirname(__file__), 'dist', triple, 'wineole-bridge.exe')
        )

    @classmethod
    def default_bridge_path(cls):
        return cls.bridge_path_for_arch(platform.machine())

    @classmethod
    def _default_spawner(cls, port):
        # wineole-bridge.exe is a native Windows binary either way -- `wine`
        # is only needed to run it on a non-Windows host. On Windows itself
        # there is no `wine` command, so running it under one would fail
        # immediately (FileNotFoundError) rather than run natively as it
        # should.
        bridge = [cls.default_bridge_path()] if os.name == 'nt' else ['wine', cls.default_bridge_path()]
        subprocess.Popen(
            bridge + [str(port)],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )

    @staticmethod
    def _default_lockfile(port):
        return os.path.join(tempfile.gettempdir(), f"wineole-bridge.{port}.lock")

    @classmethod
    def open(cls, host='127.0.0.1', port=47800, spawner=None, lockfile=None, timeout=15, token=None):
        spawner = spawner or cls._default_spawner
        lockfile = lockfile or cls._default_lockfile(port)

        sock = cls._try_connect(host, port)
        if sock:
            return cls._handshake(cls(sock), token)

        with open(lockfile, 'a+') as lock:
            cls._lock_exclusive(lock, timeout)

            sock = cls._try_connect(host, port)
            if sock:
                return cls._handshake(cls(sock), token)

            spawner(port)
            deadline = time.monotonic() + timeout
            while True:
                sock = cls._try_connect(host, port)
                if sock:
                    return cls._handshake(cls(sock), token)
                if time.monotonic() > deadline:
                    raise WineOLEError(f"wineole-bridge did not start within {timeout}s")
                time.sleep(0.2)

    @classmethod
    def _handshake(cls, client, token):
        # Mirrors wineole/client.rb's `handshake` (design doc §7.1 step 1):
        # ping after connecting, to confirm protocol compatibility and, when
        # a token is configured server-side, to present it — without this a
        # tokened bridge is unreachable from this client. A failed handshake
        # closes the socket before re-raising, so it never leaks the fd.
        try:
            client.call('ping', {'token': token} if token else {})
            return client
        except Exception:
            client.close()
            raise

    @staticmethod
    def _lock_exclusive(file_obj, timeout):
        # fcntl.flock(LOCK_EX) blocks indefinitely until the lock is free.
        # msvcrt.locking has no such mode -- LK_LOCK retries internally for
        # only about 10 seconds before raising OSError -- so on Windows this
        # wraps it in its own retry loop bounded by `timeout`, the same
        # overall budget open() already gives the whole operation.
        if os.name != 'nt':
            fcntl.flock(file_obj, fcntl.LOCK_EX)
            return

        deadline = time.monotonic() + timeout
        while True:
            try:
                file_obj.seek(0)
                msvcrt.locking(file_obj.fileno(), msvcrt.LK_LOCK, 1)
                return
            except OSError:
                if time.monotonic() > deadline:
                    raise
                time.sleep(0.2)

    @staticmethod
    def _try_connect(host, port):
        try:
            return socket.create_connection((host, port))
        except (ConnectionRefusedError, TimeoutError):
            # Only the "nothing listening (yet)" cases, mirroring the
            # ECONNREFUSED/ETIMEDOUT half of client.rb's `rescue
            # Errno::ECONNREFUSED, Errno::ETIMEDOUT, Errno::EHOSTUNREACH`.
            # Python has no dedicated OSError subclass for EHOSTUNREACH, so
            # that case isn't caught here and propagates immediately instead
            # of retrying — a deliberate, not-yet-mirrored divergence (moot
            # against the default 127.0.0.1 host). A bare `except OSError`
            # would also swallow a hostname typo (socket.gaierror) or fd
            # exhaustion (EMFILE) and retry them for the full spawn timeout
            # instead of surfacing them immediately.
            return None
