import threading
import time


class CleanupWaiters:
    """The await/signal handshake behind Client.await_cleanup and
    Client.signal_cleanup_done, pulled out of Client so it can be unit-tested
    without a live connection -- building a real Client for this would open a
    socket, and the coordination itself has nothing to do with the wire.

    One lock guards a Condition per in-flight `seq` and a map of `seq`s the
    dispatcher has already finished to the time they were finished -- the
    same shape as Mailbox's waiter, generalized from one outstanding key to
    many. Mirrors WineOLE::Client::CleanupWaiters in
    bindings/ruby/lib/wineole/client.rb.

    `_done` cannot simply be cleared by whichever `await_` consumes it: a
    signal can arrive after its `await_` already timed out, or (once
    Client.await_cleanup is in place) `await_` can be skipped entirely when
    called on the dispatcher's own thread. Either way nothing would ever
    remove that entry. Instead `signal` ages the map itself: an entry older
    than TIMEOUT can no longer have a live waiter -- any `await_` for it
    either already saw it or already gave up -- so `signal` prunes those
    entries before recording the new one, and the map never grows without
    bound even though entries are not always consumed.
    """

    # How long await_ will wait for a `seq` that never gets signalled before
    # giving up and returning anyway. A caller stuck here forever because a
    # bridge or a dispatcher died mid-cleanup would be a worse failure than
    # one that eventually gets control back, even if the instance's fate at
    # that point is unknown.
    TIMEOUT = 30.0

    def __init__(self, clock=time.monotonic):
        self._lock = threading.Lock()
        self._conds = {}
        self._done = {}
        self._clock = clock

    def await_(self, seq, timeout=None):
        """Block until signal(seq) is called from another thread, or TIMEOUT
        seconds pass, whichever comes first. Returns immediately, without
        waiting at all, when `seq` was already signalled before this call
        started.

        Named with the trailing underscore because `await` is a Python
        keyword; the Ruby method it mirrors is `CleanupWaiters#await`.
        """
        limit = self.TIMEOUT if timeout is None else timeout
        with self._lock:
            cond = self._conds.get(seq)
            if cond is None:
                # Built on this object's own lock, so `cond.wait` releases
                # exactly the lock this block holds and `signal` can get in.
                cond = threading.Condition(self._lock)
                self._conds[seq] = cond
            deadline = self._clock() + limit
            while seq not in self._done:
                left = deadline - self._clock()
                if left <= 0:
                    # A timeout returns silently: the bridge runs the steps
                    # regardless, and blocking forever on a lost ack is the
                    # worse outcome.
                    break
                cond.wait(left)
            # Cleared on the way out so a `seq` (sequence numbers are not
            # reused) never accumulates an entry once nobody can still be
            # waiting on it.
            self._done.pop(seq, None)
            self._conds.pop(seq, None)

    def signal(self, seq):
        """Record that `seq` is finished and wake whoever is waiting for it."""
        with self._lock:
            now = self._clock()
            # A seq signalled more than TIMEOUT ago can no longer have a live
            # waiter: await_ either already saw it or already gave up. Prune
            # those before recording the new one so a seq nobody ever awaits
            # does not sit in `_done` forever.
            for stale in [s for s, signalled_at in self._done.items()
                          if now - signalled_at > self.TIMEOUT]:
                del self._done[stale]
            self._done[seq] = now
            cond = self._conds.get(seq)
            if cond is not None:
                cond.notify_all()
