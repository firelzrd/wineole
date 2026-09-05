import queue
import threading
import time
import weakref

from .errors import ProtocolError, WineOLEError, warn


class _IdleSentinel:
    """Pushed on the queue when the last target or cleanup goes away."""

    __slots__ = ()

    def __repr__(self):
        return '<wineole IDLE>'


IDLE = _IdleSentinel()


class _Resume:
    """Pushed on the queue in the same critical section that puts the sink back
    on the connection, carrying the number of the era it opens. It is the FENCE
    between two eras of the queue: what is in front of it was minted while
    nothing was registered, what is behind it was minted for the registrations
    that are up now.

    Without it a frame that arrived while the connection was retired -- queued
    behind the IDLE marker, with no target in existence -- would be routed by
    the live loop to an Events that attached afterwards, because the thread
    that resumed could not tell the two eras apart.

    It carries an era rather than being a bare sentinel because a parked thread
    can miss any number of retire/re-arm cycles: with two of them behind it the
    queue holds IDLE RESUME(2) frame IDLE RESUME(3), and a fence that stopped
    at the FIRST marker it met would leave the middle era's frame in front of
    the live loop and hand it to the era-3 target. A resuming thread reads the
    current era under the lock and fences to THAT marker, so every era but the
    one it is resuming into is drained."""

    __slots__ = ('era',)

    def __init__(self, era):
        self.era = era

    def __repr__(self):
        return f'<wineole RESUME {self.era}>'


class _EndedSentinel:
    """What _drain_batches RETURNS when a batch held the end-of-stream None.
    Never put on the queue -- the two markers above are queue items, this one
    is an answer between two methods, and it exists only so that "the stream
    ended" cannot be mistaken for anything an `on_empty` or an `interrupted`
    could hand back."""

    __slots__ = ()

    def __repr__(self):
        return '<wineole ENDED>'


_ENDED = _EndedSentinel()


def _make_sink(dispatcher):
    """The connection's one sink, built by a module-level function so it
    captures a WEAK reference to the Dispatcher and NOTHING else -- a closure
    written inside a method would capture `self`, and a running reader thread
    holding it would pin the Dispatcher and, through it, the Client. The queue
    is not captured either: it is reached through the Dispatcher, which adds
    no strong edge the weak reference does not already avoid.

    Registered once, at the first attach, and does no filtering: every frame
    on the connection goes on the one queue and is routed by handle on the
    dispatcher thread, which is what makes "in arrival order, one at a time"
    true across objects rather than within one.

    This runs on the reader thread -- or, for a client whose stream has
    already ended, on the thread registering right now. It must enqueue and
    return either way: it may not block and it may not raise. A None frame
    (the end of the stream) is handed over as-is and ends the loop.

    It hands the frame to Dispatcher._enqueue rather than putting it on the
    queue itself, because whether anything is still there to take it off again
    is a question only the dispatcher's lock can answer; see _enqueue.
    """
    dispatcher_ref = weakref.ref(dispatcher)

    def sink(frame):
        dispatcher = dispatcher_ref()
        # A collected Dispatcher has no thread, no targets and no client to
        # release through: there is nothing this frame could still reach.
        if dispatcher is not None:
            dispatcher._enqueue(frame)

    return sink


class Dispatcher:
    """The one dispatcher thread of a CONNECTION, and the queue it parks on.

    One thread per connection is a promise the README makes to the caller
    ("Callbacks run on one dispatcher thread per connection, in arrival order,
    one at a time"), and the whole value of it is that a caller never needs a
    lock BETWEEN callbacks: a dict shared by an Application callback and a
    Workbook callback is safe because the two can never be inside their
    callbacks at the same time.

    THE ONE-THREAD INVARIANT, which everything below is written to keep:
    `_thread` is cleared under `_lock` ONLY when the queue is empty at that
    same moment -- on BOTH ways out of the thread, the drain that found
    nothing left (_drain_or_resume) and the end of the stream (_end_of_stream),
    which drains the same way before it clears anything. So whoever holds the
    lock and finds `_thread` set may rely on that thread to consume whatever it
    puts on the queue, and whoever finds it None starts one. There is never a
    second thread started while the first is still dispatching -- a thread that
    has read the IDLE marker does not leave, it drains; if a sink went back up
    while it was draining it simply returns to the live loop.

    A thread's life is therefore: live -> (IDLE) draining -> either live again
    (re-armed) or gone (the queue was empty under the lock).

    THE TWO MARKERS, and why the queue needs both. IDLE goes on in the same
    critical section that takes the sink OFF (nothing arrives behind it through
    the sink); RESUME goes on in the one that puts a sink BACK. Between them
    lies everything that was minted while the connection had no registrations
    at all, and none of it may be delivered: an Events that attached afterwards
    must never be handed a frame older than its own subscription. So a thread
    that finds the sink back does not just return to the live loop -- it first
    consumes the queue up to and including the RESUME, releasing what it finds
    there. Because the install and the RESUME share one critical section, a
    thread that sees `_sink` set is guaranteed the RESUME is already queued.

    A RESUME NAMES ITS ERA, because a parked thread can miss any number of
    retire/re-arm cycles rather than just one: `_arm` bumps `_era` under the
    lock and queues RESUME(era), and a resuming thread reads the era that is
    current under the lock and fences to THAT marker. Anything earlier -- an
    IDLE from a cycle it slept through, the RESUME of an era whose Events has
    since detached, and every frame minted in between -- is consumed as a
    leftover on the way. Fencing to the first marker met instead would leave
    the middle era's frames in front of the live loop, which would route them
    to the target of the era it just resumed into.

    THE ONE RESIDUAL THE MARKERS DO NOT FENCE, accepted on purpose. They only
    separate eras in which the connection had NO registrations at all, because
    only then is there a retire to hang them off. So a frame minted while
    Events A was attached, still queued when A detaches while other targets
    remain -- no retire, no IDLE, no RESUME -- and met by the live loop after
    an Events B has attached for the SAME handle, is delivered to B. B is
    handed one frame minted just before its own subscription; the frame is
    still released exactly once, so nothing leaks on the bridge. Closing it
    would mean stamping every frame with an era at mint time and carrying that
    through the queue, to narrow a window that needs one object's events to be
    unsubscribed and re-subscribed with a frame in flight between the two.
    WineOLE::Dispatcher in the Ruby client behaves the same way here.

    The one overlap the invariant permits is a thread that has ALREADY cleared
    the slot on an empty queue and is doing nothing but unwinding out of _run:
    a frame arriving in that window finds the slot empty and starts its own
    thread, so for a moment two threads exist. Only one of them can ever
    deliver anything -- the leaving one takes nothing more off the queue and
    runs no callback after the clear -- which is the promise the README makes.
    Closing even that window would mean joining the leaving thread under the
    lock, on a path the reader thread and the dispatcher itself both reach.

    Everything here is DERIVED from there being a registered callback
    somewhere on the connection, the same way Events derives its subscription
    from one: the thread and the sink go up at the first attach and come back
    down after the last detach. An ole_events nobody registered on costs no
    thread.

    THE LOCK IS NEVER HELD WHILE USER CODE RUNS. A callback that registers a
    callback on ANOTHER object reaches attach on this very thread, and would
    deadlock against a lock its own delivery still held.

    Mirrors WineOLE::Dispatcher in bindings/ruby/lib/wineole/dispatcher.rb.
    """

    def __init__(self, client):
        # Weakly, and on purpose: this object is built by the Client and held
        # by it, and an attached Events holds the Client back through
        # _targets. That ring is collected as a ring -- but only if no thread
        # holds a strong reference to any of it across a park, which is why
        # the thread body below is a staticmethod holding a weakref.
        self._client_ref = weakref.ref(client)
        # handle -> Events, at most ONE Events per handle per connection.
        # Held STRONGLY and on purpose: this is what keeps an Events alive
        # exactly as long as it has registrations, even when the caller kept
        # no reference to the Proxy it came from. Hold them weakly and the
        # Events is collected out from under a live connection: callbacks stop
        # firing, the bridge stays advised, and every event it goes on sending
        # leaks its argument handles because nobody is left to release them.
        self._targets = {}
        # handle -> on_cleanup closure. A client that uses on_cleanup but
        # subscribes to no events still needs the sink and thread up, so this
        # counts toward "is there a registered callback on the connection" the
        # same way _targets does.
        self._cleanups = {}
        # Reentrant for ONE path, and only that one: _arm calls
        # Client.on_event under this lock, and a client whose stream has
        # already ended hands the end-of-stream None straight back to the sink
        # on this very thread, which comes back in through _enqueue. Every
        # other holder of this lock is the short critical section it looks
        # like.
        self._lock = threading.RLock()
        self._queue = queue.SimpleQueue()
        self._sink = None
        # How many times a sink has been installed on this connection. Bumped
        # under the lock by _arm, which stamps the number on the RESUME marker
        # it queues, so a thread that slept through several retire/re-arm
        # cycles can tell the marker of the era it is resuming into from the
        # markers of the eras it missed. See THE TWO MARKERS above.
        self._era = 0
        # The connection's dispatcher thread, or None. See the one-thread
        # invariant above: cleared only with an empty queue.
        self._thread = None
        # The last thread ever started, kept after it has been cleared from
        # _thread and never cleared itself. For the _stopped test helper ONLY,
        # which has to tell "the slot is empty" from "and the thread that was
        # in it has really finished": the slot goes empty a moment before the
        # thread unwinds out of _run. Not consulted by on_thread -- a thread
        # running a callback of any kind, drained leftovers included, is still
        # in _thread, because the slot is only ever cleared with an empty
        # queue.
        self._last_thread = None

    def attach(self, handle, events):
        """Register `events` as the target for its handle, and put up whatever
        the connection does not have yet."""
        with self._lock:
            current = self._targets.get(handle)
            if current is not None and current is not events:
                # Only a DIFFERENT object is refused: the caller holding two
                # Events for one object. A second Events silently unseating
                # the first would leave a registered callback that never
                # fires, and either object's last off would then stop the
                # other. The same object is let through so a re-attach can
                # never be mistaken for that.
                raise ValueError(
                    f"handle {handle} already has an Events on this connection; one object's "
                    'events belong to one Events (Proxy.ole_events memoises for this reason)'
                )
            self._targets[handle] = events
            self._arm()
        return self

    def detach(self, handle, events):
        """The other half of attach, for one object. The connection's thread
        and sink only come down with the LAST one: an Application whose
        callbacks are all removed must not stop the Workbook's events on the
        same connection."""
        with self._lock:
            # By identity: removing a routing entry that belongs to somebody
            # else silently stops a live consumer.
            if self._targets.get(handle) is events:
                del self._targets[handle]
            self._release_sink_if_idle()
        return self

    def register_cleanup(self, handle, fn):
        """Register a client closure to run when the bridge asks (a $cleanup
        frame for `handle`). Arms the connection's sink and thread the same
        way attach does, so a client that uses on_cleanup but subscribes to no
        events still has a dispatcher to deliver the frame on."""
        with self._lock:
            self._cleanups[handle] = fn
            self._arm()
        return self

    def unregister_cleanup(self, handle):
        """The other half of register_cleanup. Brings the sink and thread down
        with the same idle hand-off detach's last-target case uses -- but only
        when nothing is left to deliver to: a cleanup removed while events are
        still registered must not stop them, and vice versa."""
        with self._lock:
            self._cleanups.pop(handle, None)
            self._release_sink_if_idle()
        return self

    def on_thread(self):
        """Is the calling thread this connection's own callback thread? Used
        by Client.await_cleanup to avoid a self-wait when ole_release is
        called from inside a callback -- the $cleanup frame that release
        triggers is queued behind the very callback asking the question, so
        waiting for it there would deadlock the dispatcher against itself.

        A DRAINING thread counts too, and `_thread` alone is enough to say so:
        the slot is cleared only with an empty queue, so a thread that is
        running anything at all -- a live callback, a drained $cleanup closure,
        the last leftovers before the stream ends -- is still the thread in the
        slot. A release from inside any of those would otherwise wait on the
        very thread that owes it the answer. Only the current thread can ever
        match, so this can only say "yes" about a thread that is demonstrably
        running."""
        current = threading.current_thread()
        with self._lock:
            return current is self._thread

    # --- test-only helpers -------------------------------------------------
    # Production code never needs any of these, because callbacks are the
    # delivery mechanism.

    def _stopped(self, seconds=5):
        """True once this dispatcher has no thread and the last one it started
        has really finished. Both halves matter: the slot going empty says no
        more work will be picked up, and the thread being dead says no
        callback is still running on it.

        Called ON the dispatcher thread this raises at once instead of waiting:
        a thread cannot watch itself finish, so the wait could only ever burn
        the whole budget and then report a false "still running". That is a
        misuse of the helper, and saying so immediately is the only useful
        answer."""
        deadline = time.monotonic() + seconds
        current = threading.current_thread()
        while True:
            with self._lock:
                thread = self._thread
                last = self._last_thread
            if current is thread or current is last:
                raise RuntimeError(
                    '_stopped() was called from the dispatcher thread itself, which can '
                    'never observe its own exit'
                )
            if thread is None and (last is None or not last.is_alive()):
                return True
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return False
            waiting = thread if thread is not None else last
            if waiting is not None:
                # Bounded, and re-checked: joining a LIVE dispatcher parked on
                # an empty queue would otherwise burn the whole budget in one
                # go, and the slot may be cleared by a thread that lives on
                # for a moment afterwards.
                waiting.join(min(remaining, 0.05))
            else:
                time.sleep(min(remaining, 0.005))

    def _dispatcher_thread(self):
        """The connection's dispatcher thread, live or draining."""
        with self._lock:
            return self._thread

    def _drain(self, seconds=5):
        """Block until the dispatcher has finished everything queued so far.
        Bounded, because the dispatcher not finishing is exactly what a test
        using this is hunting: an unbounded wait would hang the suite instead
        of reporting it."""
        done = threading.Event()
        # Through _enqueue like every frame, so that a barrier put after the
        # dispatcher retired is still answered instead of sitting on a queue
        # nobody reads until this call gives up. _enqueue leaves a thread
        # behind on every path it takes -- armed or retired -- and every way
        # out of a thread clears the slot only with an empty queue, so a
        # barrier on the queue always has a thread that will reach it. There
        # is no "nobody will ever call this" case left to answer inline.
        self._enqueue(('barrier', done.set))
        if not done.wait(seconds):
            raise WineOLEError(f"the dispatcher did not drain within {seconds}s")
        return self

    # --- internals ---------------------------------------------------------

    def _arm(self):
        """Put up the connection's sink and thread. Call it under the lock.

        A thread is started only when there is none. A thread that is draining
        still occupies the slot and re-arms itself the moment it sees the sink
        back (see _drain_or_resume), so this can never be the second thread.

        Putting the sink back bumps the era and puts a RESUME marker stamped
        with it on the queue in the SAME critical section, before
        Client.on_event is called -- so a thread that reads the era under the
        lock and then fences to that marker is looking for one that is already
        queued, whatever it slept through. Two things follow,
        and both are load-bearing. A thread that observes `self._sink` set is
        guaranteed the marker is already queued, so it can consume up to it
        without ever blocking. And an end-of-stream None handed straight back
        by on_event below lands BEHIND the marker, never in front of it, so a
        fence-drain can never meet the end of the stream.

        The calls out of this class made under the lock are Client.on_event
        here and Client.off_event in _release_sink_if_idle, and both are safe
        by inspection: neither runs user code (the sink is ours and only
        enqueues), and the one thing either can do on this thread -- on_event
        handing an end-of-stream None straight back when the stream has
        already ended -- comes back in through _enqueue, which is why the lock
        is reentrant and why that path only puts None on the queue (the sink
        is installed by the line above before on_event is called, so _enqueue
        sees an armed dispatcher and does nothing else with it)."""
        client = self._client_ref()
        if client is None:
            # Nothing is left to deliver to or through.
            return
        if self._sink is None:
            self._sink = _make_sink(self)
            self._era += 1
            self._queue.put(_Resume(self._era))
            client.on_event(self._sink)
        if self._thread is None:
            self._start_thread()

    def _start_thread(self):
        """Give the connection its dispatcher thread. Call it under the lock,
        with `self._thread` None: the lock is what makes "start one only if
        there is none" safe against a second arm and against the reader.

        A staticmethod and a weakref, never a bound method or a closure over
        self: a running thread is a strong reference to whatever its target
        holds, so a thread started from a method here would be a permanent GC
        root for this Dispatcher -- and through it for the Client, whose
        socket would then never be closed."""
        thread = threading.Thread(
            target=Dispatcher._run,
            args=(weakref.ref(self), self._queue),
            name='wineole-dispatcher',
            daemon=True,
        )
        self._thread = thread
        self._last_thread = thread
        thread.start()

    def _enqueue(self, item):
        """One frame from the sink, on the reader thread -- or, for a stream
        that has already ended, on the thread registering right now. THE ONLY
        WAY ANYTHING GETS ON THE QUEUE.

        Under the lock, because the question "is anybody still going to take
        this off the queue" is decided by the same critical sections that
        retire the sink (_release_sink_if_idle), clear the thread slot
        (_drain_or_resume) and end the stream (_end_of_stream). Serialised
        against those, a frame the reader was already carrying when the sink
        came off can no longer fall between them.

        Taking the lock on the reader thread cannot deadlock against the
        reader: this class never holds the lock while waiting on it. Nothing
        under the lock makes a call (no client.call anywhere near it), and the
        only calls out of the class under it -- Client.on_event and
        Client.off_event -- take the Client's sink lock and the Mailbox lock
        briefly and never wait for a response. The one case that comes
        straight back here on the same thread, on_event handing an
        end-of-stream None back from inside _arm, is why the lock is
        reentrant.
        """
        with self._lock:
            if self._sink is not None:
                # Armed: _arm leaves a thread behind whenever it leaves a sink
                # behind, and the invariant says that thread is still there to
                # consume this.
                self._queue.put(item)
                if self._thread is None:
                    # Armed with no thread, which is ORDINARY rather than an
                    # anomaly: every first _arm of a closed connection reaches
                    # it. _arm installs the sink and calls Client.on_event
                    # before it starts the thread, and on an already-ended
                    # stream that call hands the end-of-stream None straight
                    # back to the sink on this very thread -- so the None
                    # re-enters here with a sink up and the slot still empty,
                    # and it is this branch, not _arm, that starts the thread
                    # which eats it. (A thread killed by an unexpected
                    # BaseException, whose finally clears the slot, leaves the
                    # same state.) Starting one here is what keeps either from
                    # silently stranding every later frame on the queue (see
                    # the retired branch for the cost of a start here).
                    self._start_thread()
                return
            if item is None:
                # The stream ended after retirement: nothing to release, and
                # no consumer left that would need telling to stop.
                return
            # Retired: the sink is off. This frame would otherwise sit on the
            # queue forever -- its argument handles never released, and, for a
            # $cleanup, a thread parked in Client.await_cleanup stalled for the
            # full timeout.
            start = self._thread is None
            if start:
                # Nobody is left, so this thread has to make one, and the IDLE
                # marker goes on AHEAD of the item, never behind it. Ahead, it
                # sends the new thread into the drain before the thread can
                # reach the item, so the item is always consumed as a leftover
                # (released, or, for a $cleanup, run and acked) and never by
                # the live loop -- and, with an empty queue, the drain takes
                # the new thread out again.
                #
                # The order is what makes that true whatever the timing. The
                # new thread's first `get` takes no lock, so an _arm landing
                # between the start below and that get cannot be ordered from
                # outside; with the item in FRONT, such an _arm would leave the
                # live loop looking at a frame minted while the connection had
                # no registrations at all and hand it to the Events that just
                # attached. Behind the marker it cannot be reached that way: a
                # thread that finds the sink back fences up to the RESUME the
                # _arm queued, and the item is in front of that.
                self._queue.put(IDLE)
            self._queue.put(item)
            if start:
                # Thread.start waits for the new thread to begin, and this is
                # on the reader thread: a bounded cost, paid only in this race
                # window, and the alternative is a frame whose handles are
                # never released.
                self._start_thread()
            # A thread that IS there has not passed an empty check yet -- that
            # check clears the slot, under this lock -- so it will drain this.

    def _release_sink_if_idle(self):
        """Take the sink off the connection once nothing is left to deliver
        to, and hand the thread its notice. Call it under the lock."""
        if self._targets or self._cleanups:
            return
        sink = self._sink
        if sink is None:
            return
        self._sink = None
        client = self._client_ref()
        if client is not None:
            client.off_event(sink)
        # After the sink comes off, never before, and inside the same critical
        # section every enqueue has to take: the marker means "nothing more
        # arrives behind this through the sink", and _enqueue is what makes
        # that true rather than merely likely. A reader already inside the
        # sink -- it read the sink out of the Mailbox before off_event removed
        # it -- blocks in _enqueue until this section ends and then takes the
        # retired path there, so its frame lands behind this marker and is
        # drained by the same thread. Never a join, either -- the thread
        # reaching here is very often the dispatcher itself, in a callback
        # that called off.
        self._queue.put(IDLE)

    @staticmethod
    def _run(dispatcher_ref, work_queue):
        """The dispatcher thread's body, on the class so it captures no
        instance. The Dispatcher is reached weakly and never held across a
        get, which is what lets a Client with events on it still be collected.

        THIS THREAD SURVIVES EVERYTHING. A dead dispatcher is permanent and
        silent: the callbacks stay registered, the bridge stays advised, every
        later event's argument handles leak for the life of the connection,
        the queue grows without bound, and on_error is never told. So the
        except is BaseException, not Exception: the narrow one exists to let a
        fatal exception reach a thread that can act on it, and there is no
        such thread here.

        Four items are the loop's own business rather than a callback's: the
        end-of-stream None ends it, the IDLE marker sends it into the drain
        (from which it either comes back live or leaves), a RESUME marker met
        HERE is ignored whatever era it names, and everything else is delivery.

        A RESUME reaches the live loop when an _arm did not follow a retire --
        the first arm of the connection, most often -- so there is no era
        boundary to enforce and nothing in front of it to fence off. The one a
        resuming thread does care about it consumes in _fence_to_resume,
        before it ever gets back here.
        """
        dispatcher = None
        try:
            while True:
                item = work_queue.get()
                dispatcher = dispatcher_ref()
                if dispatcher is None:
                    return  # there is nobody left to deliver to
                if item is None:
                    dispatcher._end_of_stream()
                    return
                if isinstance(item, _Resume):
                    # Nothing to do with it here -- but drop the strong
                    # reference before parking again, for the same reason the
                    # bottom of the loop does.
                    dispatcher = None
                    continue
                if item is IDLE:
                    try:
                        resume = dispatcher._drain_or_resume()
                    except BaseException as exc:  # noqa: BLE001
                        dispatcher._report(exc, item)
                        resume = False
                    if not resume:
                        return
                else:
                    try:
                        dispatcher._handle(item)
                    except BaseException as exc:  # noqa: BLE001 -- see above
                        dispatcher._report(exc, item)
                # Dropped before parking again: a local still pointing at the
                # Dispatcher would pin it for as long as this thread waits,
                # which is precisely what the weakref is here to avoid.
                dispatcher = None
        finally:
            # Only ever a backstop: every ordinary way out of the loop clears
            # the slot itself, under the lock and with an empty queue. If this
            # thread died some other way the Dispatcher must still not go on
            # believing it has one, or nothing would start a successor.
            dispatcher = dispatcher_ref()
            if dispatcher is not None:
                dispatcher._thread_finished(threading.current_thread())

    @staticmethod
    def _is_barrier(item):
        """A _drain barrier: ('barrier', callable). Recognised in one place so
        the live loop and the leftover drain cannot come to disagree about
        what one looks like."""
        return isinstance(item, tuple) and bool(item) and item[0] == 'barrier'

    def _handle(self, item):
        """One queued item on the LIVE path: the connection is armed, so a
        frame goes to the object it names."""
        if self._is_barrier(item):
            item[1]()
            return
        if isinstance(item, dict):
            # A $cleanup frame goes to the client's on_cleanup closure and is
            # acked; every other frame is routed by handle to its Events.
            if item.get('event') == '$cleanup':
                self._run_cleanup(item)
            else:
                self._route(item)
            return
        # Everything else names nothing to deliver to and holds nothing to
        # release.

    def _handle_leftover(self, item):
        """One item taken off the queue by a DRAINING thread: the sink was off
        the connection when this was drained, so there was no target for it at
        that moment.

        A $cleanup is still run and acked -- it owes the bridge a
        release_event and owes a thread in Client.await_cleanup its wake-up,
        and its closure was registered before the retirement, not after. Any
        other frame is RELEASED, never routed: at drain time no target existed
        for it, and an Events that attaches afterwards must not be handed a
        frame minted before its own subscription. Releasing is what keeps the
        COM objects behind those handles from sitting on the bridge until the
        connection closes."""
        if self._is_barrier(item):
            # A barrier drained rather than met on the live path must still be
            # answered, or whoever is waiting on it hangs.
            item[1]()
            return
        if isinstance(item, dict):
            if item.get('event') == '$cleanup':
                self._run_cleanup(item)
            else:
                self._release(item)
            return
        # A marker met here -- an IDLE from an earlier retire/re-arm cycle, a
        # RESUME in front of the one a fence-drain is looking for -- has
        # nothing to deliver and nothing to release. (The end-of-stream None is
        # taken out of the batch by the caller before it reaches here.)

    def _drain_or_resume(self):
        """What the dispatcher thread does when it reaches an IDLE marker.
        True to go back to the live loop, False to end the thread.

        The slot is cleared -- and the thread therefore ends -- ONLY with an
        empty queue under the lock, which is the whole one-thread invariant:
        a frame enqueued a moment earlier is still on the queue and is drained
        on the next pass, and a frame enqueued a moment later finds the slot
        empty and starts a thread of its own. There is no gap between the two.

        A sink back on the connection wins outright: whoever re-armed did so
        after this marker was queued, so this thread goes back to live
        delivery -- but only after it has consumed the queue up to the RESUME
        of the era that is current at that moment, because everything in front
        of that marker was minted for registrations that no longer exist."""

        def re_armed():
            # Asked under the lock before every pass. A sink back on the
            # connection means this thread is not draining any more: what is
            # queued now was minted for the registrations that are up, and
            # taking it into a batch would release it instead of delivering it.
            #
            # The ERA is read in this same critical section and returned, so
            # the fence below looks for the marker of the era that was current
            # when the drain stopped -- not merely the first marker it meets,
            # which may belong to a cycle this thread slept through. It is
            # always >= 1 when a sink is up (_arm bumps it before installing
            # one), so it is never a falsey "keep draining".
            return self._era if self._sink is not None else False

        def retire():
            # Under the lock, with the queue empty at this very moment: the
            # only place this thread may clear the slot.
            self._thread = None
            return False

        outcome = self._drain_batches(retire, interrupted=re_armed)
        if outcome is _ENDED:
            # The end of the stream, drained rather than met on the live path.
            self._end_of_stream()
            return False
        if outcome:
            # Re-armed while this thread was working. Back to live delivery on
            # the SAME thread -- no successor was started, because the slot was
            # never cleared -- once every era before the one `outcome` names
            # has been fenced off. Outside the lock: what it finds there runs
            # client closures.
            self._fence_to_resume(outcome)
            return True
        return False

    def _drain_batches(self, on_empty, interrupted=None):
        """Under the lock take EVERYTHING the queue holds; outside the lock
        handle that batch as leftovers; repeat until a pass finds nothing left.
        Both ways out of the live loop are this one loop -- the IDLE drain
        (_drain_or_resume) and the end of the stream (_end_of_stream) -- and
        they differ only in what happens on empty.

        `on_empty` is called UNDER the lock, at the moment the queue is empty,
        and what it returns is returned from here. That is the only place a
        caller may clear the thread slot, and the empty check sharing one
        critical section with the clear is the whole one-thread invariant: a
        frame enqueued a moment earlier is still in a batch, a frame enqueued
        a moment later finds the slot empty and starts a thread of its own,
        and there is no gap between the two.

        `interrupted`, if given, is asked under the lock BEFORE each pass
        whether the drain should stop; anything truthy it returns is returned
        from here, for the caller to act on OUTSIDE the lock (what a resume
        does next runs client closures). Before the pass and not after: a batch
        taken here is drained, never delivered.

        A None in a batch is the end of the stream, drained rather than met on
        the live path. The rest of the batch is still handled -- those items
        were minted before the stream ended -- and then _ENDED is returned, for
        the caller to say what the end of the stream means to it."""
        while True:
            with self._lock:
                if interrupted is not None:
                    stop = interrupted()
                    if stop:
                        return stop
                batch = []
                while True:
                    try:
                        batch.append(self._queue.get_nowait())
                    except queue.Empty:
                        break
                if not batch:
                    return on_empty()
            # Outside the lock: a leftover $cleanup runs a client closure, and
            # a closure is free to touch anything on this connection.
            ended = False
            for item in batch:
                if item is None:
                    ended = True
                    continue
                try:
                    self._handle_leftover(item)
                except BaseException as exc:  # noqa: BLE001
                    # The item that failed, never the marker that drained it:
                    # _report routes by handle, so reporting the marker would
                    # send the failure of a frame past the object it belongs to
                    # and its on_error.
                    self._report(exc, item)
            if ended:
                return _ENDED

    def _fence_to_resume(self, era):
        """Consume the queue up to AND INCLUDING the RESUME marker of `era`,
        handling what is in front of it as leftovers. Call it WITHOUT the lock:
        a leftover $cleanup runs a client closure.

        This is the era boundary. Everything in front of that marker was minted
        while the connection was retired, or for registrations of an era that
        has since ended -- there was no target for it by the time this thread
        got here, and the Events that re-armed attached afterwards, so
        delivering it would hand a subscriber a frame older than its own
        subscription. It is released instead, which is also what keeps the COM
        objects behind those handles from sitting on the bridge until the
        connection closes.

        The marker is matched by era, never by being the first one met: a
        thread parked in a callback can sleep through several retire/re-arm
        cycles, and stopping at an earlier marker would leave the frames of the
        eras in between in front of the live loop, to be routed to the target
        of the era this thread is resuming into. Those earlier markers are
        consumed and ignored like any other leftover.

        `get()` without a timeout, on purpose: _arm bumps the era and puts its
        RESUME on the queue in the same critical section in which it installs
        the sink, so a thread that read `era` under that same lock is looking
        at a queue that already holds the marker. It cannot block for longer
        than it takes to handle what is in front of it."""
        while True:
            item = self._queue.get()
            if isinstance(item, _Resume) and item.era == era:
                return
            if item is None:
                # Not reachable: _arm queues the RESUME before it installs the
                # sink, and a retired _enqueue drops the end-of-stream None
                # entirely, so no None can be in front of a marker this thread
                # is looking for. Put it back rather than swallow it, so that
                # if the ordering were ever broken the stream would still end
                # instead of leaving a thread parked forever on an empty queue.
                self._queue.put(item)
                return
            try:
                self._handle_leftover(item)
            except BaseException as exc:  # noqa: BLE001 -- see _drain_or_resume
                self._report(exc, item)

    def _end_of_stream(self):
        """The stream is over: no frame will ever arrive again. What is still
        on the queue is handled first, and both slots go only once there is
        nothing left -- literally the same loop as the drain in
        _drain_or_resume, _drain_batches, differing only in what it does once
        the queue is empty. Clearing the thread slot over a queue that still
        held items is what once stranded a $cleanup's ack and a _drain barrier
        on a queue nobody would ever read again.

        The sink comes OFF the connection rather than merely being forgotten:
        the Client holds the strong reference to it, so dropping it here alone
        left one behind in Client._event_sinks for every later _arm on the
        closed connection to add to.

        A later _arm installs a fresh sink; on an already-closed Mailbox that
        sink is handed the end-of-stream None straight back, and the thread
        _arm starts eats it and arrives here again -- which is why this has to
        be repeatable and has to leave nothing behind."""

        def close_down():
            # Under the lock, with the queue empty at this very moment: see
            # _drain_batches for why both slots may only go from there.
            sink = self._sink
            self._sink = None
            if sink is not None:
                client = self._client_ref()
                if client is not None:
                    client.off_event(sink)
            self._thread = None

        # No `interrupted`: past the end of the stream there is no era to go
        # back to, so this drain runs to the end whatever is armed behind it --
        # and takes that sink off too.
        while self._drain_batches(close_down) is _ENDED:
            # A second end-of-stream marker, met in a batch. The stream can
            # only end once; there is nothing more to do about it here, but
            # what was queued behind it still has to be drained.
            pass

    def _thread_finished(self, thread):
        with self._lock:
            if self._thread is thread:
                self._thread = None

    def _run_cleanup(self, frame):
        """A $cleanup frame: run the client closure for its handle on THIS
        thread (COM-safe, like every other callback), then tell the bridge the
        closure is done and wake whoever is blocked in await_cleanup. The
        closure's own exception must not stop any of that -- the bridge runs
        the steps regardless, so a raising closure still ends with
        release_event and the waiter signalled. Nor may a MISSING closure:
        a $cleanup whose registration is already gone (drained behind the
        marker that retired it, say) still owes the bridge its ack and still
        owes a waiting thread its wake-up.

        The closure is looked up UNDER the lock and called WITHOUT it: like a
        callback, it is free to touch anything on this connection, each of
        which re-enters this class and would deadlock against a lock its own
        delivery still held."""
        handle = frame.get('handle')
        seq = frame.get('seq')
        with self._lock:
            fn = self._cleanups.get(handle)
        try:
            if fn is not None:
                fn()
        except Exception as exc:
            warn(f"wineole: on_cleanup raised {type(exc).__name__}: {exc}")
        finally:
            self._ack_cleanup(handle, seq)

    def _ack_cleanup(self, handle, seq):
        """The three steps every $cleanup ends with, whether a closure was
        registered for it or not, whether that closure ran or raised, and
        whether the step before this one raised: tell the bridge the frame is
        finished, wake whatever thread is parked in Client.await_cleanup for
        it, and drop the registration -- which is also what brings the sink
        and the thread down once nothing is left.

        Guarded step by step rather than written as a sequence: a raise out of
        any one of them used to skip the rest, and the two things skipped are
        exactly the ones nobody else will do -- a waiter that is never
        signalled parks for the full CleanupWaiters.TIMEOUT, and a
        registration that is never dropped keeps the connection's thread and
        sink up forever. BaseException, not Exception, for the same reason the
        dispatcher loop catches it: there is no thread above this one that
        could act on it."""
        client = self._client_ref()
        if client is not None:
            try:
                client.call('release_event', {'seq': seq})
            except (ProtocolError, OSError):
                # The connection is going away; there is nothing left to
                # release. Expected, so not worth a line on stderr. Only those
                # two: Client.call raises ProtocolError for both ways a
                # connection ends under it (the stream already over, and the
                # socket failing mid-write, which it wraps), and OSError covers
                # a socket failure raised outside that wrapping. Anything else
                # -- a RemoteError from a bridge that is still very much there,
                # a bug in this class -- reaches the reporter below and is
                # warned rather than swallowed.
                pass
            except BaseException as exc:  # noqa: BLE001
                warn(f"wineole: cleanup {seq} release_event raised {type(exc).__name__}: {exc}")
            try:
                client.signal_cleanup_done(seq)
            except BaseException as exc:  # noqa: BLE001
                warn(f"wineole: cleanup {seq} signal raised {type(exc).__name__}: {exc}")
        try:
            self.unregister_cleanup(handle)
        except BaseException as exc:  # noqa: BLE001
            warn(f"wineole: cleanup {seq} unregister raised {type(exc).__name__}: {exc}")

    def _route(self, frame):
        """One frame, to the object it names. The target is looked up under
        the lock and called WITHOUT it: delivery runs user code, and the
        callback is free to register or remove anything on this connection,
        which comes straight back here as an attach or a detach."""
        try:
            target = self._target_for(frame)
            # A frame for a handle with no target was minted before the
            # unsubscribe reached the bridge. Not delivered -- off means off
            # -- but still released below, because the COM objects behind
            # those handles would otherwise sit on the bridge until the
            # connection closed.
            if target is not None:
                target._deliver(frame)
        finally:
            # The arguments are valid for the callback and no longer.
            # Releasing here is what makes that true even when the frame never
            # reached a callback at all.
            self._release(frame)

    def _release(self, frame):
        """The frames of one event are released together, whether they reached
        a callback, reached one that raised, reached no target at all, or were
        still queued when the last target left."""
        # args: null says the bridge minted NOTHING for this event: it is
        # serialized as null rather than left out precisely so the client can
        # tell that from "this event had zero arguments", and nothing is
        # inserted in the bridge's event table for it. Sending a release
        # anyway is a synchronous round trip from the dispatcher, which caps
        # the event rate at one RTT -- in exactly the args=False case a caller
        # reaches for to keep up with a high-frequency event.
        if not isinstance(frame, dict) or frame.get('args') is None:
            return
        client = self._client_ref()
        if client is None:
            return
        try:
            client.call('release_event', {'seq': frame.get('seq')})
        except (ProtocolError, OSError):
            # The connection is going away; there is nothing left to release.
            # Narrow on purpose, exactly as in _ack_cleanup: anything else
            # propagates to the dispatcher loop's per-item reporter, which
            # warns about it instead of losing it here.
            pass

    def _target_for(self, item):
        if not isinstance(item, dict):
            return None
        with self._lock:
            return self._targets.get(item.get('handle'))

    def _report(self, exc, item):
        """Never raises. It is called from the dispatcher's own except, so an
        exception out of here is the one thing that could still end the
        thread. Routed to the object the frame names, so that an on_error
        registered on it is told; a frame that names nobody has no handler to
        reach."""
        try:
            target = self._target_for(item)
            if target is not None:
                target._report(exc, item)
            else:
                warn(f"wineole: dispatcher raised {type(exc).__name__}: {exc}")
        except BaseException:  # noqa: BLE001
            pass  # even stderr being gone must not end the dispatcher
