import ipaddress
import itertools
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

from .errors import WineOLEError, ProtocolError, RemoteError, InstanceClosingError
from .mailbox import Mailbox, CLOSED
from .proxy import Proxy
from .cleanup_waiters import CleanupWaiters
from .dispatcher import Dispatcher

_client_generation = itertools.count(1)

# Mirrors wineole/client.rb's ARCH_TRIPLES (Ruby, Task 5) -- kept in sync by
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
        self.generation = next(_client_generation)
        # The one strong reference to the event sinks; see on_event. Written
        # there and never read for delivery -- the reader thread reaches the
        # sinks through the Mailbox's weak references instead, so this list is
        # what keeps them alive. Delete it as dead code and a plain
        # client.on_event(...) keeps working only until the next garbage
        # collection, after which the connection silently stops delivering
        # events. The guard is ClientEventsTest.
        # test_a_registered_consumer_survives_a_garbage_collection.
        self._event_sinks = []
        self._event_sinks_lock = threading.Lock()
        # Eagerly, and deliberately: a Dispatcher is two dicts, a queue and
        # two empty slots until something attaches to it -- no thread, no
        # sink, no socket traffic -- so building it here costs less than the
        # lock that building it lazily would need, on a connection that may
        # have several objects registering callbacks at once.
        self._dispatcher = Dispatcher(self)
        self._mailbox = Mailbox(sock)
        # Backs await_cleanup/signal_cleanup_done.
        self._cleanup_waiters = CleanupWaiters()

    def __del__(self):
        # Safety net for a forgotten close() -- best-effort, not a
        # replacement for it (GC timing is never guaranteed). Exceptions must
        # not escape __del__: the interpreter would just print them to stderr
        # and continue, so there's nothing to gain by letting one propagate.
        try:
            self.close()
        except Exception:
            pass

    def on_event(self, sink):
        """Register a consumer of server-initiated frames (those with no
        `id`). `sink` is called INLINE on the reader thread, which is the only
        thread reading this socket: it must neither block nor raise. A sink
        that waited would stall every response on the connection; one that
        made a COM call of its own would leave nobody to read the answer and
        deadlock against itself. Enqueue and return.

        It is also called once with None when the stream ends, so a consumer
        parked on a queue can finish instead of blocking on it forever --
        including when it registers after the stream has already ended, in
        which case it is handed None right here.

        APPENDS, never replaces. A replacing registration would silently
        switch an earlier consumer off. The events feature puts up exactly one
        sink here for the whole connection (Dispatcher) and routes by handle
        on the dispatcher thread, but nothing about this method assumes that.

        The strong reference is held HERE rather than in the Mailbox, because
        the running reader thread pins the Mailbox: a sink almost always
        closes over the object that owns it, and that object almost always
        holds this Client, so a strong reference from the Mailbox would pin
        the Client too -- no Client could be collected, __del__ would never
        run and its socket would stay open for the life of the process.

        `sink` must be weak-referenceable: a plain function or a closure
        works, but a builtin such as `print` does not and raises here.
        `off_event` matches by identity, so the caller must pass back the
        exact object it registered here -- a fresh bound method created for
        the call to `off_event` is a different object and will not match."""
        with self._event_sinks_lock:
            self._event_sinks.append(sink)
        self._mailbox.add_sink(sink)

    def off_event(self, sink):
        """The way back out: a sink registered for the life of the connection
        is a consumer that cannot be dismantled. Identity, not equality -- two
        consumers can compare equal without being the same registration, and
        removing the wrong one silently stops a live consumer."""
        with self._event_sinks_lock:
            self._event_sinks = [s for s in self._event_sinks if s is not sink]
        self._mailbox.remove_sink(sink)

    @property
    def dispatcher(self):
        """This connection's ONE dispatcher: the thread every callback on it
        runs on, in arrival order, one at a time. Per connection rather than
        per object because that is the promise the README makes to a caller
        who shares state between an Application callback and a Workbook
        callback -- they can never be inside their callbacks at the same time,
        so no lock of their own is needed.

        This Client and its Dispatcher hold each other, and an attached Events
        holds this Client back through the Dispatcher's target table. That
        ring is collected as a ring: no thread ever holds a strong reference
        to any of it across a park, so the whole connection is still
        collectible with callbacks registered on it -- which is what __del__
        needs to be true if it is ever to close the socket."""
        return self._dispatcher

    def await_cleanup(self, seq):
        """Block until the dispatcher finishes the $cleanup for `seq` (the
        client closure, then the release_event that follows it). If the caller
        IS the dispatcher thread -- ole_release called from inside a callback
        -- do not wait: the $cleanup frame is queued behind the current
        callback and will only run after it returns, so waiting here would
        deadlock the dispatcher against itself."""
        if self.on_dispatcher_thread():
            return
        self._cleanup_waiters.await_(seq)

    def signal_cleanup_done(self, seq):
        """Called by the dispatcher once it has finished delivering $cleanup
        `seq`, to release whatever thread is parked in await_cleanup for it."""
        self._cleanup_waiters.signal(seq)

    def on_dispatcher_thread(self):
        """Is the calling thread this connection's own dispatcher thread?"""
        return self._dispatcher.on_thread()

    def call(self, method, params=None):
        try:
            response = self._mailbox.request(
                {'method': method, 'params': params if params is not None else {}}
            )
        except OSError as exc:
            # The socket went away mid-write. Reported as this client's own
            # protocol error, with the OSError kept as the cause.
            raise ProtocolError('connection closed') from exc
        # A sentinel, not a frame: the end of the stream is something that
        # happened here, and saying so in the shape of a wire error would make
        # it indistinguishable from one the bridge really sent.
        if response is CLOSED:
            raise ProtocolError('connection closed')
        if response.get('error'):
            klass = response['error']['class']
            # WineOLE::InstanceClosingError is the one remote error class this
            # client resolves to its own local class rather than wrapping in a
            # generic RemoteError -- so a caller can catch it directly instead
            # of pattern-matching on RemoteError.remote_class. Mirrors
            # client.rb's `call`.
            if klass == 'WineOLE::InstanceClosingError':
                raise InstanceClosingError(response['error']['message'])
            raise RemoteError(klass, response['error']['message'])
        return response.get('result')

    def close(self):
        # Ends the stream: every pending caller is woken, every consumer is
        # told, the reader thread finishes and the fd is closed (the buffered
        # reader holds an io-ref on the socket, so closing the socket alone
        # would not let the peer see EOF -- which is what triggers the
        # bridge's session/handle reclamation). It does NOT release live
        # roots; that is ole_release's job.
        self._mailbox.close()
        # Told directly, through a normal strong attribute, and not only
        # through the sink the Mailbox just tried to notify. When this Client
        # is being closed by its own __del__ as part of collecting the
        # Client<->Dispatcher<->Events cycle, CPython has already cleared
        # every weak reference INTO that cycle before running any finalizer
        # in it -- the sink's weakref back to this Dispatcher included -- so
        # the Mailbox's attempt to hand the sink its end-of-stream None can
        # silently do nothing, and the dispatcher thread would stay parked on
        # its queue forever. `self._dispatcher` is a plain attribute of this
        # still-alive `self`, never a weakref, so it is unaffected and always
        # gets the connection's dispatcher to notice the stream has ended.
        # Harmless to call when the dispatcher was never armed (it is a no-op
        # then) and harmless if the Mailbox's own delivery already succeeded
        # (a second end-of-stream None past the first is ignored).
        self._dispatcher._enqueue(None)

    @property
    def loopback(self):
        # Mirrors client.rb's `loopback?`: is the bridge on the other end of
        # this connection reachable only through the loopback interface --
        # i.e. the same machine?
        #
        # Deliberately the same test the bridge itself uses to decide whether
        # a token is required (`peer_addr().ip().is_loopback()` in main.rs).
        # Anything that keys off "is this local" -- path conversion in the
        # Office wrapper, for one -- must agree with the bridge, or a
        # connection ends up remote for authentication and local for
        # everything else.
        #
        # `getpeername()[0]` is the IP string for both address families --
        # IPv4 returns (host, port), IPv6 returns (host, port, flowinfo,
        # scopeid), but index 0 is the host in either shape. `is_loopback`
        # covers all of 127.0.0.0/8 and ::1, matching Rust's
        # `IpAddr::is_loopback`. A host's own NIC address is NOT loopback,
        # and that is intended: it is a different machine as far as this
        # boundary cares.
        try:
            return ipaddress.ip_address(self._sock.getpeername()[0]).is_loopback
        except Exception:
            return False

    def create(self, class_name, cleanup=None):
        return Proxy.create(class_name, self, cleanup=cleanup)

    def connect(self, class_name, cleanup=None):
        return Proxy.connect(class_name, self, cleanup=cleanup)

    def connect_or_create(self, class_name, cleanup=None):
        return Proxy.connect_or_create(class_name, self, cleanup=cleanup)

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
        # Mirrors wineole/client.rb's `handshake` (design doc section 7.1
        # step 1): ping after connecting, to confirm protocol compatibility
        # and, when a token is configured server-side, to present it --
        # without this a tokened bridge is unreachable from this client. A
        # failed handshake closes the socket before re-raising, so it never
        # leaks the fd.
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
            sock = socket.create_connection((host, port))
            # Mirrors client.rb's try_connect and the bridge's own
            # set_nodelay. Insurance only: the ~40 ms per-RPC stall this
            # project hit was on the response side, and requests already go
            # out in a single write, so this changes nothing today -- it
            # keeps a future multi-write request path from reintroducing it.
            sock.setsockopt(socket.IPPROTO_TCP, socket.TCP_NODELAY, 1)
            return sock
        except (ConnectionRefusedError, TimeoutError):
            # Only the "nothing listening (yet)" cases, mirroring the
            # ECONNREFUSED/ETIMEDOUT half of client.rb's `rescue
            # Errno::ECONNREFUSED, Errno::ETIMEDOUT, Errno::EHOSTUNREACH`.
            # Python has no dedicated OSError subclass for EHOSTUNREACH, so
            # that case isn't caught here and propagates immediately instead
            # of retrying -- a deliberate, not-yet-mirrored divergence (moot
            # against the default 127.0.0.1 host). A bare `except OSError`
            # would also swallow a hostname typo (socket.gaierror) or fd
            # exhaustion (EMFILE) and retry them for the full spawn timeout
            # instead of surfacing them immediately.
            return None
