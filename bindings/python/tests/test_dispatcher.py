import contextlib
import io
import os
import sys
import threading
import time
import unittest

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from fake_bridge import FakeBridge
from wineole.client import Client
from wineole.dispatcher import Dispatcher


# The Dispatcher's $cleanup delivery: a $cleanup frame runs the client's
# on_cleanup closure ON THE DISPATCHER THREAD (the COM-safe context, like
# every other callback), then acks the bridge with release_event(seq) and
# wakes whoever is blocked in await_cleanup. The closure's own exception must
# not stop any of that.
class FakeCleanupClient:
    """Shaped like the real Client for exactly the calls the Dispatcher makes
    on the cleanup path: call / on_event / off_event / signal_cleanup_done."""

    def __init__(self):
        self.calls = []
        self.signalled = []
        self.sink = None
        # Set on the dispatcher thread once the ack has been sent, so a test
        # waits on evidence rather than on a clock.
        self.done = threading.Event()

    def call(self, method, params=None):
        self.calls.append((method, params))
        return None

    def on_event(self, sink):
        self.sink = sink

    def off_event(self, sink):
        if self.sink is sink:
            self.sink = None

    def signal_cleanup_done(self, seq):
        self.signalled.append(seq)
        self.done.set()


class DispatcherTest(unittest.TestCase):
    def test_cleanup_frame_runs_proc_on_dispatcher_thread_and_acks(self):
        client = FakeCleanupClient()
        ran_on = []
        answered = []

        dispatcher = Dispatcher(client)

        def closure():
            ran_on.append(threading.current_thread())
            # The question Client.await_cleanup asks before it waits: a
            # release called from inside a callback must not wait for a
            # $cleanup queued behind the very callback asking, which would
            # deadlock the dispatcher against itself.
            answered.append(dispatcher.on_thread())

        dispatcher.register_cleanup(7, closure)
        # Feed a $cleanup frame the way the reader would.
        client.sink({'event': '$cleanup', 'handle': 7, 'seq': 99, 'args': None})

        self.assertTrue(client.done.wait(5), 'the cleanup must be acked within 5s')
        self.assertEqual(len(ran_on), 1, 'the closure runs exactly once')
        self.assertIsNot(ran_on[0], threading.current_thread(),
                         'the closure must not run on the thread that pushed the frame')
        self.assertEqual(ran_on[0].name, 'wineole-dispatcher',
                         'it must run on the connection dispatcher thread, the COM-safe context')
        self.assertEqual(answered, [True],
                         'on_thread must answer True from the dispatcher thread itself')
        self.assertTrue(dispatcher._stopped(5),
                        'the last registration going away takes the thread down with it')
        self.assertIn(('release_event', {'seq': 99}), client.calls)
        self.assertEqual(client.signalled, [99])

    def test_cleanup_still_acks_when_closure_raises(self):
        client = FakeCleanupClient()

        def boom():
            raise RuntimeError('boom')

        dispatcher = Dispatcher(client)
        dispatcher.register_cleanup(7, boom)
        # A raising closure warns on stderr on purpose; keep it out of the
        # run's output.
        with contextlib.redirect_stderr(io.StringIO()):
            client.sink({'event': '$cleanup', 'handle': 7, 'seq': 5, 'args': None})
            self.assertTrue(client.done.wait(5), 'the cleanup must be acked within 5s')

        self.assertIn(('release_event', {'seq': 5}), client.calls)
        self.assertEqual(client.signalled, [5],
                         'the bridge runs the steps regardless, so a raising closure still acks')
        self.assertTrue(dispatcher._stopped(5),
                        'the last registration going away takes the thread down with it, '
                        'closure raised or not')


class RecordingTarget:
    """The smallest thing the Dispatcher will route to: an Events stand-in
    that records what it was handed. Task 5 brings the real _deliver."""

    def __init__(self):
        self.delivered = []
        self.reported = []

    def _deliver(self, frame):
        self.delivered.append(frame)

    def _report(self, exc, frame):
        self.reported.append((exc, frame))


class ConnectedDispatcherMixin:
    """One real Client on one real socket, for the tests below that assert
    what reaches the BRIDGE (release_event) and what wakes a thread in
    await_cleanup as well as what reaches a target -- none of which a stand-in
    for the Client would exercise. Shared by every such test class, so the two
    of them cannot drift into two different set-ups of the same thing."""

    @contextlib.contextmanager
    def connected(self):
        with FakeBridge(handler=lambda method, params: None) as bridge:
            client = Client(bridge.sock)
            try:
                yield bridge, client
            finally:
                client.close()


class RetiredDispatcherTest(ConnectedDispatcherMixin, unittest.TestCase):
    """The window between the sink coming off the connection and the thread
    draining what is left behind it: the reader can already be inside the sink
    with a frame when the sink is removed, so that frame arrives at a
    dispatcher that has nothing left to consume it. It must still be handled --
    a $cleanup acked, any other frame released -- rather than left on the
    queue, where its argument handles would sit on the bridge for the life of
    the connection and a thread in Client.await_cleanup would stall for the
    full timeout.
    """

    def retired_dispatcher(self, client):
        """A dispatcher armed and brought back down: its sink is off the
        connection and its thread has exited. That is exactly the state a
        frame the reader was already carrying finds."""
        dispatcher = client.dispatcher
        dispatcher.register_cleanup(1, lambda: None)
        dispatcher.unregister_cleanup(1)
        self.assertTrue(dispatcher._stopped(5),
                        'the last registration going away takes the thread down with it')
        return dispatcher

    def test_a_cleanup_frame_left_behind_the_idle_marker_is_still_acked(self):
        with self.connected() as (bridge, client):
            dispatcher = self.retired_dispatcher(client)
            # The reader, already inside the sink when the sink came off.
            dispatcher._enqueue({'event': '$cleanup', 'handle': 1, 'seq': 7})
            started = time.monotonic()
            client.await_cleanup(7)
            self.assertLess(time.monotonic() - started, 1.0,
                            'await_cleanup must be woken by the ack, not time out')
            self.assertIn(('release_event', {'seq': 7}), bridge.requests,
                          'a $cleanup arriving after retirement is still acked to the bridge')
            self.assertTrue(dispatcher._stopped(5),
                            'the thread started to handle it exits again')

    def test_a_frame_arriving_after_the_sink_retired_is_released_not_left_on_the_queue(self):
        with self.connected() as (bridge, client):
            dispatcher = self.retired_dispatcher(client)
            dispatcher._enqueue({'event': 'X', 'handle': 1, 'seq': 9, 'args': []})
            self.assertTrue(dispatcher._stopped(5),
                            'the thread started to handle it exits again')
            self.assertIn(('release_event', {'seq': 9}), bridge.requests,
                          'a frame nobody is left to deliver to still has its handles released')

    def test_end_of_stream_after_retirement_does_nothing(self):
        with self.connected() as (bridge, client):
            dispatcher = self.retired_dispatcher(client)
            # Every dispatcher thread alive right now, including any another
            # test in this process is still winding down: the assertion below
            # is about a thread THIS call would start, not about the process
            # being quiet.
            before = {t for t in threading.enumerate() if t.name == 'wineole-dispatcher'}
            dispatcher._enqueue(None)
            self.assertTrue(dispatcher._stopped(1),
                            'the end of the stream after retirement leaves no thread running')
            after = {t for t in threading.enumerate() if t.name == 'wineole-dispatcher'}
            self.assertEqual(after - before, set(),
                             'there is nothing to release and no consumer to end: no thread')


class OneDispatcherThreadTest(ConnectedDispatcherMixin, unittest.TestCase):
    """A connection has ONE dispatcher thread at a time. The interesting
    moment is the one after the last unsubscribe, when the thread has read its
    IDLE marker and is working through what was left on the queue: it has not
    gone anywhere, so nothing may start a second thread beside it, and what it
    drains was minted before that unsubscribe, so nothing may be delivered to
    an Events that attached afterwards.
    """

    @staticmethod
    def gate():
        """A closure that parks the dispatcher thread wherever it is called,
        so a test can arrange the queue behind it instead of racing it."""
        entered = threading.Event()
        release = threading.Event()

        def run():
            entered.set()
            if not release.wait(10):
                raise AssertionError('the dispatcher gate was never released')

        return run, entered, release

    def hold_the_thread_and_retire(self, dispatcher, seq=1):
        """Park the dispatcher thread inside a cleanup closure and retire the
        connection under it. The IDLE marker is queued while the thread cannot
        reach it, so whatever the test enqueues next lands behind that marker
        and is drained rather than delivered."""
        run, entered, release = self.gate()
        dispatcher.register_cleanup(1, run)
        dispatcher._enqueue({'event': '$cleanup', 'handle': 1, 'seq': seq})
        self.assertTrue(entered.wait(5), 'the dispatcher must reach the closure')
        dispatcher.unregister_cleanup(1)
        return release

    def test_only_one_dispatcher_thread_ever_runs(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            # Threads another test in this process is still winding down are
            # not this test's business.
            others = {t for t in threading.enumerate() if t.name == 'wineole-dispatcher'}
            release_closure = self.hold_the_thread_and_retire(dispatcher, seq=11)

            # A barrier behind the IDLE marker parks the thread INSIDE the
            # drain -- the window a second thread used to be started in.
            drain_gate, in_drain, release_drain = self.gate()
            dispatcher._enqueue(('barrier', drain_gate))
            release_closure.set()
            self.assertTrue(in_drain.wait(5), 'the thread must reach the drained barrier')

            # Two frames arriving at a retired connection while its one thread
            # is mid-drain. It is still in the slot, so it takes them.
            dispatcher._enqueue({'event': 'X', 'handle': 1, 'seq': 12, 'args': []})
            dispatcher._enqueue({'event': 'X', 'handle': 1, 'seq': 13, 'args': []})

            peak = [0]

            def sample():
                live = [t for t in threading.enumerate()
                        if t.name == 'wineole-dispatcher' and t not in others]
                peak[0] = max(peak[0], len(live))

            for _ in range(500):
                sample()
            release_drain.set()
            deadline = time.monotonic() + 5
            while time.monotonic() < deadline:
                sample()
                if dispatcher._dispatcher_thread() is None:
                    break
            self.assertTrue(dispatcher._stopped(5),
                            'an empty queue under the lock takes the thread down')
            sample()

            self.assertEqual(peak[0], 1,
                             'one dispatcher thread per connection, at every moment')
            self.assertIn(('release_event', {'seq': 12}), bridge.requests)
            self.assertIn(('release_event', {'seq': 13}), bridge.requests)

    def test_a_leftover_frame_is_released_not_delivered_to_a_later_events(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            release_closure = self.hold_the_thread_and_retire(dispatcher, seq=20)
            drain_gate, in_drain, release_drain = self.gate()
            dispatcher._enqueue(('barrier', drain_gate))
            # Minted before the last unsubscribe: it is behind the IDLE marker.
            dispatcher._enqueue({'event': 'X', 'handle': 5, 'seq': 21, 'args': []})
            release_closure.set()
            self.assertTrue(in_drain.wait(5), 'the thread must reach the drained barrier')

            # An Events attaches for that very handle while the thread is
            # mid-drain -- the race that used to hand it a stale frame.
            target = RecordingTarget()
            dispatcher.attach(5, target)
            release_drain.set()

            dispatcher._drain(5)
            self.assertEqual(target.delivered, [],
                             'a frame minted before a subscription is never delivered to it')
            self.assertIn(('release_event', {'seq': 21}), bridge.requests,
                          'it is released instead, so its handles leave the bridge')

            # The thread saw the sink back and returned to the live loop
            # instead of exiting, so a fresh frame reaches the target.
            fresh = {'event': 'X', 'handle': 5, 'seq': 22, 'args': []}
            dispatcher._enqueue(fresh)
            dispatcher._drain(5)
            self.assertEqual(target.delivered, [fresh])
            self.assertIn(('release_event', {'seq': 22}), bridge.requests)

            dispatcher.detach(5, target)
            self.assertTrue(dispatcher._stopped(5))

    def test_draining_thread_resumes_when_re_armed(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            release_closure = self.hold_the_thread_and_retire(dispatcher, seq=30)
            thread = dispatcher._dispatcher_thread()
            self.assertIsNotNone(thread, 'the thread is held in the closure, not gone')

            # Re-armed before the thread ever reaches its IDLE marker.
            target = RecordingTarget()
            dispatcher.attach(5, target)
            fresh = {'event': 'X', 'handle': 5, 'seq': 31, 'args': []}
            dispatcher._enqueue(fresh)
            release_closure.set()

            dispatcher._drain(5)
            self.assertEqual(target.delivered, [fresh],
                             'what is queued after a re-arm is routed, not drained')
            self.assertIs(dispatcher._dispatcher_thread(), thread,
                          'the same thread carried on; a re-arm starts no second one')
            dispatcher.detach(5, target)
            self.assertTrue(dispatcher._stopped(5))

    def test_end_of_stream_drains_what_is_queued_behind_it(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            run, entered, release = self.gate()
            dispatcher.register_cleanup(1, run)
            dispatcher._enqueue({'event': '$cleanup', 'handle': 1, 'seq': 40})
            self.assertTrue(entered.wait(5), 'the dispatcher must reach the closure')

            # The stream ends under a thread that cannot reach it yet, so the
            # end-of-stream None sits in FRONT of whatever the test queues now.
            client.close()

            outcome = []

            def drain():
                try:
                    dispatcher._drain(2)
                    outcome.append('returned')
                except BaseException as exc:  # noqa: BLE001 -- reported below
                    outcome.append(exc)

            drainer = threading.Thread(target=drain, name='test-drainer')
            drainer.start()
            try:
                # The barrier has to be BEHIND the None for this test to be
                # about anything, so wait until it is really on the queue
                # rather than assuming the drainer thread got there first.
                deadline = time.monotonic() + 5
                while dispatcher._queue.qsize() < 2 and time.monotonic() < deadline:
                    time.sleep(0.001)
                self.assertGreaterEqual(
                    dispatcher._queue.qsize(), 2,
                    'the barrier must be queued behind the end-of-stream None')
                release.set()
            finally:
                drainer.join(10)
            self.assertEqual(outcome, ['returned'],
                             'the end of the stream drains what is queued behind it, so a '
                             'barrier there is answered instead of stranded')
            self.assertTrue(dispatcher._stopped(5),
                            'and the thread leaves once that queue really is empty')

    def test_a_frame_queued_before_a_re_arm_is_released_not_delivered(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            release_closure = self.hold_the_thread_and_retire(dispatcher, seq=50)
            # Minted before the re-arm: it is behind the IDLE marker the
            # retirement left, and no target existed for it at that moment.
            dispatcher._enqueue({'event': 'X', 'handle': 5, 'seq': 51, 'args': []})
            # The Events attaches BEFORE the thread ever reaches that marker.
            # The frame is now in front of a live sink, which is exactly the
            # ordering that used to hand it to a target minted after it.
            target = RecordingTarget()
            dispatcher.attach(5, target)
            release_closure.set()

            dispatcher._drain(5)
            self.assertEqual(target.delivered, [],
                             'a frame queued before a re-arm is never delivered to the '
                             'Events that re-armed')
            self.assertIn(('release_event', {'seq': 51}), bridge.requests,
                          'it is released instead, so its handles leave the bridge')

            # The thread crossed the marker the re-arm left and went back to
            # the live loop, so a frame minted AFTER the attach is delivered.
            fresh = {'event': 'X', 'handle': 5, 'seq': 52, 'args': []}
            dispatcher._enqueue(fresh)
            dispatcher._drain(5)
            self.assertEqual(target.delivered, [fresh],
                             'what is queued after the re-arm is routed on the same thread')
            self.assertIn(('release_event', {'seq': 52}), bridge.requests)

            dispatcher.detach(5, target)
            self.assertTrue(dispatcher._stopped(5))

    def test_two_retire_re_arm_cycles_before_the_thread_wakes_are_both_fenced(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            release_closure = self.hold_the_thread_and_retire(dispatcher, seq=70)

            # Two whole retire/re-arm cycles while the thread is parked, so
            # the queue holds two eras' worth of markers at once:
            #   IDLE(1) RESUME(2) frame_a IDLE(2) RESUME(3) frame_b
            # A fence that stopped at the FIRST resume marker it met would
            # cross only IDLE(1) and RESUME(2), and hand frame_a -- minted for
            # an Events that has since detached -- to the live loop, which
            # would route it to the Events that attached in era 3.
            first = RecordingTarget()
            dispatcher.attach(5, first)
            frame_a = {'event': 'X', 'handle': 5, 'seq': 71, 'args': []}
            dispatcher._enqueue(frame_a)
            dispatcher.detach(5, first)

            second = RecordingTarget()
            dispatcher.attach(5, second)
            frame_b = {'event': 'X', 'handle': 5, 'seq': 72, 'args': []}
            dispatcher._enqueue(frame_b)

            release_closure.set()
            dispatcher._drain(5)

            self.assertEqual(first.delivered, [],
                             'the Events of the middle era detached before the thread ever '
                             'reached its frame')
            self.assertEqual(second.delivered, [frame_b],
                             'frame_a belongs to an era the fence must cross; only the frame '
                             'queued after the last re-arm is delivered live')
            self.assertIn(('release_event', {'seq': 71}), bridge.requests,
                          'the fenced-off frame is released instead, so its handles leave '
                          'the bridge')
            self.assertIn(('release_event', {'seq': 72}), bridge.requests)

            dispatcher.detach(5, second)
            self.assertTrue(dispatcher._stopped(5))

    def test_a_frame_enqueued_while_fully_retired_is_released_even_if_an_events_attaches_at_once(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            # Fully retired: no sink on the connection AND no thread left, the
            # state in which an enqueue has to start a thread of its own.
            dispatcher.register_cleanup(1, lambda: None)
            dispatcher.unregister_cleanup(1)
            self.assertTrue(dispatcher._stopped(5),
                            'the last registration going away takes the thread down with it')

            # The reader, still inside the sink when the connection retired
            # completely -- and an Events attaching for that very handle in the
            # same breath, possibly before the thread this enqueue starts has
            # taken anything off the queue at all (its first get takes no
            # lock, so the test cannot order the two from outside). It does not
            # need to: the IDLE marker goes on the queue AHEAD of the frame, so
            # the frame is behind a marker either way and can only ever be
            # consumed as a leftover -- released, never routed.
            target = RecordingTarget()
            dispatcher._enqueue({'event': 'X', 'handle': 5, 'seq': 61, 'args': []})
            dispatcher.attach(5, target)

            dispatcher._drain(5)
            self.assertEqual(target.delivered, [],
                             'a frame minted while the connection was retired is never '
                             'delivered to an Events that attaches right after it')
            self.assertIn(('release_event', {'seq': 61}), bridge.requests,
                          'it is released instead, so its handles leave the bridge')

            # The thread the enqueue started crossed the marker the attach left
            # and went to the live loop, so a frame minted AFTER the attach is
            # delivered.
            fresh = {'event': 'X', 'handle': 5, 'seq': 62, 'args': []}
            dispatcher._enqueue(fresh)
            dispatcher._drain(5)
            self.assertEqual(target.delivered, [fresh],
                             'what is minted after the attach is routed on that same thread')
            self.assertIn(('release_event', {'seq': 62}), bridge.requests)

            dispatcher.detach(5, target)
            self.assertTrue(dispatcher._stopped(5))

    def test_repeated_arms_on_a_closed_connection_do_not_accumulate_sinks(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            dispatcher.register_cleanup(1, lambda: None)
            client.close()
            self.assertTrue(dispatcher._stopped(5),
                            'the end of the stream ends the dispatcher thread')
            self.assertEqual(len(client._event_sinks), 0,
                             'the end of the stream takes the sink off the connection too')

            # Every one of these arms installs a fresh sink on a Mailbox that
            # is already closed, which hands the end-of-stream None straight
            # back. Each sink must come off again, or a long-lived Client
            # collects one per registration for the rest of its life.
            target = RecordingTarget()
            for round_ in range(3):
                dispatcher.attach(5, target)
                dispatcher.register_cleanup(2, lambda: None)
                self.assertTrue(dispatcher._stopped(5),
                                f"round {round_}: the thread the arm started leaves again")
                self.assertLessEqual(len(client._event_sinks), 1,
                                     f"round {round_}: arming a closed connection must not "
                                     'accumulate sinks')

    def test_drain_after_the_stream_ended_does_not_hang(self):
        with self.connected() as (bridge, client):
            dispatcher = client.dispatcher
            dispatcher.register_cleanup(1, lambda: None)
            client.close()
            self.assertTrue(dispatcher._stopped(5),
                            'the end of the stream ends the dispatcher thread')
            started = time.monotonic()
            # The end of the stream takes the SINK down as well as the thread.
            # Left installed, it made this barrier look like it had a consumer.
            dispatcher._drain(5)
            self.assertLess(time.monotonic() - started, 2.0,
                            'a barrier on an ended connection is answered, not timed out')
            self.assertTrue(dispatcher._stopped(5),
                            'and the thread it took to answer it leaves again')


if __name__ == '__main__':
    unittest.main()
