import itertools
import json
import socket
import threading
import weakref

from .errors import warn


class _ClosedSentinel:
    """The value a waiter is handed when the stream ends.

    An object, not a frame: the end of the stream is something that happened
    HERE, and saying so in the shape of a wire error would make it
    indistinguishable from one the bridge really sent.
    """

    __slots__ = ()

    def __repr__(self):
        return '<wineole CLOSED>'


CLOSED = _ClosedSentinel()


class _Waiter:
    """One response, handed from the reader thread to the caller."""

    __slots__ = ('_event', '_lock', '_value')

    def __init__(self):
        self._event = threading.Event()
        self._lock = threading.Lock()
        self._value = None

    def fill(self, value):
        """First write wins. The reader fills a waiter and wakes its caller,
        but the caller has not yet re-acquired the Mailbox lock to delete
        itself from the table; if the stream ends in that window, a
        last-write-wins slot would overwrite the answer the bridge really sent
        with "connection closed" and the caller would raise for a request that
        succeeded."""
        with self._lock:
            if self._value is not None:
                return
            self._value = value
            self._event.set()

    def take(self):
        self._event.wait()
        return self._value


class Mailbox:
    """The socket, the waiter table and the event sinks -- everything the
    reader thread shares with the calling threads.

    A separate object for a garbage-collection reason, not a tidiness one: a
    running reader thread is a strong reference to whatever its target is
    bound to, so a reader started as a method of Client would make every
    Client permanently reachable -- __del__ would never run and its socket
    would stay open for the life of the process. The thread's target lives
    here instead, and this object holds no reference back to the Client, not
    even through the event sinks: those it reaches only weakly (see
    Client.on_event). Mirrors WineOLE::Client::Mailbox.
    """

    def __init__(self, sock):
        self._sock = sock
        self._ids = itertools.count(1)
        # Guards the socket WRITE and the bookkeeping below -- never a wait
        # for a response. Holding a lock across the round trip is what made a
        # COM call from inside an event callback impossible to even send while
        # another thread was waiting.
        self._lock = threading.Lock()
        self._waiters = {}
        # Weak references, deliberately: the Client owns the sinks. This list
        # only lets the reader thread reach them.
        self._sinks = []
        self._closed = False
        self._reader_file = sock.makefile('rb')
        self._reader = threading.Thread(target=self._read_loop, name='wineole-reader', daemon=True)
        self._reader.start()

    def request(self, params):
        """Send one request and block this thread -- and only this thread --
        until the reader routes the matching response back. Returns the
        response frame, or CLOSED when the stream ended first."""
        request_id = None
        waiter = _Waiter()
        try:
            with self._lock:
                # A request on a dead connection is not sent at all: no reader
                # is left to route an answer, so the caller would wait on its
                # slot for the life of the process.
                if self._closed:
                    return CLOSED
                request_id = next(self._ids)
                self._waiters[request_id] = waiter
                payload = dict(params)
                payload['id'] = request_id
                line = json.dumps(payload, ensure_ascii=False) + '\n'
                self._sock.sendall(line.encode('utf-8'))
            return waiter.take()
        finally:
            if request_id is not None:
                with self._lock:
                    self._waiters.pop(request_id, None)

    def add_sink(self, sink):
        """Register a consumer of server-initiated frames. Held weakly here;
        the Client holds the strong reference."""
        with self._lock:
            closed = self._closed
            if not closed:
                self._sinks.append(weakref.ref(sink))

        # Outside the lock, because this runs user code on the CALLER's
        # thread: a consumer that closed the client from its end-of-stream
        # branch would otherwise deadlock on a lock this method still held.
        # Without this hand-off a consumer that attached after the bridge died
        # would never be told, and its dispatcher thread would park on an
        # empty queue for the life of the process.
        if closed:
            self._deliver(sink, None)

    def remove_sink(self, sink):
        """Drop one sink by identity, and every dead weak reference met on the
        way past -- the reader only prunes those when it delivers, and a quiet
        connection would otherwise keep them."""
        with self._lock:
            self._prune_sinks(drop=sink)

    def close(self):
        """Idempotent. Ends the stream and joins the reader."""
        try:
            # shutdown before anything else, and lock-free: a caller wedged
            # in sendall under _lock must not be able to block close() for
            # as long as that write is stuck. shutdown does not need the
            # lock, and closing a socket does not wake a thread already
            # blocked reading it -- the reader is the thread this has to end.
            self._sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        # Every pending caller is woken HERE rather than only from the
        # reader's own end-of-stream sweep: a reader parked inside a consumer
        # that issued a call of its own can never reach that sweep, and the
        # caller it is blocking is the reader itself. This also sets
        # _closed.
        self._fail_waiters()
        # A sink runs on the reader thread, so a close called from inside one
        # would have that thread join itself. The reader is on its way out
        # anyway: the socket is shut down, so its next read ends the loop.
        if self._reader is not threading.current_thread():
            self._reader.join(2)
        self._close_socket()

    def _close_socket(self):
        # The buffered reader holds an io-ref on the socket, so closing the
        # socket alone defers the real fd close until the reader object is
        # finalized -- the peer would not see EOF and getpeername would go on
        # answering. Close both, in this order, and only ever from here.
        try:
            self._reader_file.close()
        except OSError:
            pass
        try:
            self._sock.close()
        except OSError:
            pass

    def _read_loop(self):
        try:
            while True:
                line = self._reader_file.readline()
                if not line:
                    break
                try:
                    frame = json.loads(line)
                except ValueError:
                    # An unparseable line is skipped, not fatal.
                    continue
                # null, 123 and [] are all valid JSON and none of them is a
                # frame. Skipped exactly like an unparseable line: asking a
                # non-dict for a key would raise, which would kill the reader
                # and every later call on the connection with it.
                if not isinstance(frame, dict):
                    continue
                if 'id' in frame:
                    try:
                        with self._lock:
                            waiter = self._waiters.get(frame['id'])
                    except TypeError:
                        # An id that cannot be hashed (a list, a dict) can
                        # never match a waiter. Skipped like any other
                        # malformed frame, not fatal to the reader.
                        continue
                    if waiter is not None:
                        waiter.fill(frame)
                else:
                    # Handed off, never run here.
                    self._dispatch(frame)
        except (OSError, ValueError):
            # OSError: the socket went away. ValueError: the buffered reader
            # was closed under this thread by a consumer that called close.
            pass
        finally:
            self._fail_waiters()
            # Tell every consumer the stream is over, so a dispatcher thread
            # parked on a queue can finish instead of blocking forever.
            self._dispatch(None)
            self._close_socket()

    def _fail_waiters(self):
        """A waiter that is never woken waits forever, so the end of the
        stream has to reach every one of them."""
        with self._lock:
            self._closed = True
            pending = list(self._waiters.values())
            self._waiters.clear()
        for waiter in pending:
            waiter.fill(CLOSED)

    def _dispatch(self, frame):
        # The sinks are copied out under the lock and called WITHOUT it.
        # Holding it across the dispatch would make a call issued from inside
        # a consumer unable to even reach the wire -- the reader would be
        # holding the very lock that guards the socket write, on the very
        # thread the consumer runs on.
        for sink in self._live_sinks():
            self._deliver(sink, frame)

    def _live_sinks(self):
        with self._lock:
            return self._prune_sinks()

    def _prune_sinks(self, drop=None):
        """Walk self._sinks once, dropping every dead weak reference and,
        if given, the one matching `drop` by identity. Returns the sinks
        that remain live, in registration order. Caller must hold _lock."""
        live = []
        kept = []
        for ref in self._sinks:
            found = ref()
            if found is None or found is drop:
                continue
            kept.append(ref)
            live.append(found)
        self._sinks = kept
        return live

    @staticmethod
    def _deliver(sink, frame):
        """One misbehaving consumer must not take the connection down with it.
        The sinks share a single reader thread, so an exception raised out of
        one of them would end the read loop -- every other consumer would stop
        seeing events and every later call would fail with "connection
        closed". Reported rather than swallowed: a sink that raises is a bug
        in the sink."""
        try:
            sink(frame)
        except Exception as exc:
            warn(f"wineole: event consumer raised {type(exc).__name__}: {exc}")
