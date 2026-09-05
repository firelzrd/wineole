import os
import sys
import threading
import time
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.cleanup_waiters import CleanupWaiters


# The await/signal handshake behind Client.await_cleanup and
# Client.signal_cleanup_done, tested in isolation from Client itself:
# building a real Client for this would open a socket, and the coordination
# has nothing to do with the wire.
class CleanupWaitersTest(unittest.TestCase):
    def test_signal_from_another_thread_unblocks_a_waiting_thread(self):
        waiters = CleanupWaiters()
        returned = []
        # Set by the waiter thread immediately before it calls await_, so the
        # assertion below waits on evidence that the thread really got there
        # rather than on a clock -- this test is about the notify waking an
        # ALREADY-waiting thread, not about the already-signalled case the
        # next test covers.
        about_to_wait = threading.Event()

        def wait_for_it():
            about_to_wait.set()
            waiters.await_(7)
            returned.append(True)

        waiter = threading.Thread(target=wait_for_it)
        waiter.start()

        self.assertTrue(about_to_wait.wait(5), 'the waiter thread must reach await_')
        self.assertEqual(returned, [], 'must still be blocked before signal is sent')

        waiters.signal(7)

        waiter.join(5)
        self.assertFalse(waiter.is_alive(), 'await_ must return once another thread calls signal')
        self.assertEqual(returned, [True])

    def test_await_returns_immediately_when_already_signalled(self):
        waiters = CleanupWaiters()
        waiters.signal(9)

        # Bounded well under CleanupWaiters.TIMEOUT (30s): if await_ actually
        # had to wait for that, this fails instead of merely taking longer
        # than it should.
        started = time.monotonic()
        waiters.await_(9)
        self.assertLess(time.monotonic() - started, 1.0,
                        'an already-signalled seq must not wait at all')

    def test_different_seqs_do_not_share_state(self):
        waiters = CleanupWaiters()
        waiters.signal(1)

        # Signalling seq 1 must not leak into seq 2: a thread awaiting a
        # different, unsignalled seq blocks until it is signalled on its own.
        returned = []
        # Set immediately before await_, so this waits on the waiter thread
        # having really reached it rather than on a clock.
        about_to_wait = threading.Event()

        def wait_for_two():
            about_to_wait.set()
            waiters.await_(2)
            returned.append(True)

        waiter = threading.Thread(target=wait_for_two)
        waiter.start()
        self.assertTrue(about_to_wait.wait(5), 'the waiter thread must reach await_')
        self.assertEqual(returned, [], 'seq 2 must not be affected by seq 1 having been signalled')

        waiters.signal(2)
        waiter.join(5)
        self.assertFalse(waiter.is_alive())
        self.assertEqual(returned, [True])

    def test_a_signal_nobody_awaits_is_forgotten_after_the_timeout(self):
        # A seq that gets signalled but never awaited (a timed-out await_, or
        # one skipped entirely) must not sit in _done forever: signal() ages
        # the map itself, pruning entries older than TIMEOUT. Reading _done
        # directly rather than adding a _done_seqs_for_test() helper -- it is
        # already a plain dict and this is the only test that needs to look.
        now = [0.0]
        waiters = CleanupWaiters(clock=lambda: now[0])

        waiters.signal(1)
        self.assertEqual(set(waiters._done.keys()), {1})

        now[0] = CleanupWaiters.TIMEOUT + 1
        waiters.signal(2)

        self.assertEqual(set(waiters._done.keys()), {2},
                         'seq 1 must be pruned once it is older than TIMEOUT')

        started = time.monotonic()
        waiters.await_(2)
        self.assertLess(time.monotonic() - started, 1.0,
                        'pruning must not eat a signal still fresh enough to matter')


if __name__ == '__main__':
    unittest.main()
