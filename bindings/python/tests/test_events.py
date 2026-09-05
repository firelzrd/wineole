import contextlib
import datetime
import gc
import io
import itertools
import json
import os
import queue
import socket
import sys
import threading
import unittest
import weakref

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.dispatcher import Dispatcher
from wineole.errors import RemoteError
from wineole.events import Events, Subscription
from wineole.proxy import Proxy

sys.path.insert(0, os.path.dirname(__file__))

from fake_bridge import FakeBridge
from wineole.client import Client

_fake_generation = itertools.count(1)


def waited(box, seconds=5):
    """A bounded get. Under every bug these tests name -- a callback that
    never runs, an error that never reaches on_error, a frame delivered to the
    wrong Events -- nothing is ever put, so an unbounded get would hang the
    suite forever instead of failing it. Timing out returns None, which fails
    the same assertion in seconds with the expected value still in the
    message."""
    try:
        return box.get(timeout=seconds)
    except queue.Empty:
        return None


def frame(event, seq, args=None, handle=1):
    return {'event': event, 'handle': handle, 'seq': seq, 'args': args}


class FakeClient:
    """Records what reached the bridge, and lets a test push frames in as if
    they had arrived on the socket."""

    def __init__(self):
        self.calls = []
        self.sinks = []
        # A fresh, distinct generation per instance, mirroring the real
        # Client -- Proxy.decode reads it when an event argument is an object
        # reference.
        self.generation = next(_fake_generation)
        self._dispatcher = None

    def call(self, method, params=None):
        self.calls.append((method, params))
        return True

    @property
    def dispatcher(self):
        # One per connection, exactly as the real Client has one -- which is
        # the property most of the tests below are about. Lazily rather than
        # in __init__ only because a fake needs no answer to the question the
        # real Client builds it eagerly for; every test here reaches it from
        # one thread first.
        if self._dispatcher is None:
            self._dispatcher = Dispatcher(self)
        return self._dispatcher

    # Appends, mirroring the real Client -- a fake that replaced the sink
    # could not reach the two-objects-on-one-connection case at all.
    def on_event(self, sink):
        self.sinks.append(sink)

    # By identity, mirroring the real Client for the same reason it does.
    def off_event(self, sink):
        self.sinks = [s for s in self.sinks if s is not sink]

    def deliver(self, frame_or_none):
        for sink in list(self.sinks):
            sink(frame_or_none)

    def count(self, method):
        return sum(1 for name, _ in self.calls if name == method)


class RefusingClient(FakeClient):
    """A client whose bridge refuses every subscribe, the way an object that
    is not an event source does."""

    def call(self, method, params=None):
        super().call(method, params)
        if method == 'subscribe':
            raise RemoteError('WIN32OLERuntimeError', 'not an event source')
        return True


class GatedClient(FakeClient):
    """A client whose call for one method blocks until the test lets it
    through. That is what makes an interleaving deterministic rather than
    hopeful: the thread under test is held INSIDE the wire call, in exactly
    the window where it has decided something and not yet carried it out."""

    def __init__(self, gated_method):
        super().__init__()
        self._gated_method = gated_method
        self._entered = queue.Queue()
        self._release = queue.Queue()

    def call(self, method, params=None):
        if method == self._gated_method:
            self._entered.put(method)
            self._release.get()
        return super().call(method, params)

    def wait_until_inside(self, seconds=5):
        return waited(self._entered, seconds)

    def let_it_through(self):
        self._release.put('go')


class OnceReleaseGatedClient(FakeClient):
    """The same idea as GatedClient for one call only: the FIRST
    release_event blocks, and every one after it goes straight through.
    One-shot on purpose -- the window this holds open is the dispatcher's idle
    hand-off, and everything the test does after it has to be able to
    complete."""

    def __init__(self):
        super().__init__()
        self._entered = queue.Queue()
        self._release = queue.Queue()
        self._already_gated = False

    def call(self, method, params=None):
        if method == 'release_event' and not self._already_gated:
            self._already_gated = True
            self._entered.put(params['seq'])
            self._release.get()
        return super().call(method, params)

    def wait_until_inside(self, seconds=5):
        return waited(self._entered, seconds)

    def let_it_through(self):
        self._release.put('go')


class EventsTest(unittest.TestCase):
    def events_for(self, client=None, handle=1):
        client = FakeClient() if client is None else client
        return Events(client, handle), client

    # L1/L2 are derived: registering a callback is the ONLY thing the caller
    # does, and the subscribe must follow from it. If it did not, the callback
    # would be registered and the event would never arrive -- silently.
    def test_registering_a_callback_subscribes_on_the_bridge(self):
        ev, client = self.events_for()

        def callback(*args):
            pass

        sub = ev.on('SheetChange', callback)
        self.assertEqual(
            client.calls,
            [('subscribe', {'handle': 1, 'event': 'SheetChange', 'args': True})])

        # What `on` hands back: an opaque, read-only registration. Read-only
        # because `off` and the wire flag are derived from what it holds, so a
        # caller rewriting one of them would change what the bridge was told
        # without telling it.
        self.assertIsInstance(sub, Subscription)
        self.assertEqual(sub.name, 'SheetChange')
        self.assertIs(sub.callback, callback)
        self.assertIs(sub.args, True)
        with self.assertRaises(AttributeError):
            sub.name = 'Other'
        # Python's TypeError stands where Ruby raises ArgumentError for a
        # missing block: there is nothing to call otherwise.
        with self.assertRaises(TypeError):
            ev.on('SheetChange', 'not callable')

    # An object that is not an event source is refused by the bridge, and that
    # refusal arrives here as an exception out of on. If the callback stayed
    # registered anyway, the caller would be left holding a Subscription for
    # an event that can never arrive -- the one state this class is built to
    # make unreachable -- and a later on for the same event would not even try
    # to subscribe again.
    def test_a_refused_subscribe_leaves_no_callback_behind(self):
        client = RefusingClient()
        ev = Events(client, 1)

        with self.assertRaises(RemoteError):
            ev.on('Click', lambda *args: None)
        with self.assertRaises(RemoteError):
            ev.on('Click', lambda *args: None)

        self.assertEqual(
            client.count('subscribe'), 2,
            'the refused registration must be gone, so the next on tries again rather than '
            'finding a callback already listed and subscribing to nothing')
        self.assertEqual(ev._registered_names(), [])
        self.assertEqual(client.sinks, [], 'the refused object must leave nothing on the connection')

    def test_a_second_callback_for_the_same_event_does_not_subscribe_again(self):
        ev, client = self.events_for()
        ev.on('SheetChange', lambda *args: None)
        ev.on('SheetChange', lambda *args: None)
        self.assertEqual(client.count('subscribe'), 1)

    def test_removing_the_last_callback_unsubscribes(self):
        ev, client = self.events_for()
        first = ev.on('SheetChange', lambda *args: None)
        ev.on('SheetChange', lambda *args: None)
        ev.off(first)
        self.assertEqual(client.count('unsubscribe'), 0, 'one left, still subscribed')
        ev.off('SheetChange')
        self.assertEqual(client.count('unsubscribe'), 1)

    # The bridge holds one args flag per event, and on is the only place the
    # caller states what it wants. Measured on Excel before this was derived:
    # a second callback asking for arguments next to an args=False one was
    # handed nothing, because the flag from the first registration was still
    # standing and nothing re-subscribed. The wire flag is the union of the
    # live callbacks, in both directions.
    def test_the_wire_flag_follows_what_the_callbacks_asked_for(self):
        ev, client = self.events_for()
        ev.on('Click', lambda *args: None, args=False)
        wants_args = ev.on('Click', lambda *args: None)
        self.assertEqual(
            client.calls,
            [('subscribe', {'handle': 1, 'event': 'Click', 'args': False}),
             ('subscribe', {'handle': 1, 'event': 'Click', 'args': True})],
            'a callback that wants the arguments must re-subscribe for them')

        ev.off(wants_args)
        self.assertEqual(
            client.calls[-1], ('subscribe', {'handle': 1, 'event': 'Click', 'args': False}),
            'and with it gone, stop paying for handles nobody asked for')
        self.assertEqual(client.count('unsubscribe'), 0,
                         'one callback is left, so the subscription itself stays')

    # after is None cannot tell "the last callback just went" from "there was
    # never one", and an unsubscribe for a subscription that does not exist is
    # a round trip that says nothing -- on an object the bridge may never have
    # advised at all.
    def test_off_for_a_name_that_was_never_registered_touches_nothing(self):
        ev, client = self.events_for()
        self.assertIsNone(ev.off('NeverRegistered'))

        self.assertEqual(client.calls, [], 'there is no subscription to take down')
        self.assertEqual(ev._registered_names(), [], 'and a read must not leave a key behind')
        self.assertEqual(client.sinks, [], 'nor put a consumer on the connection')

    # The registry is mutated under one lock and the bridge is told outside
    # it, so two threads can reach the wire in the opposite order to the one
    # they decided in. Measured on the Ruby code before its wire mutex: wire
    # order subscribe, subscribe, unsubscribe -- with a callback still
    # registered, i.e. a callback whose event can never arrive.
    def test_a_subscribe_must_not_land_after_the_unsubscribe_that_follows_it(self):
        client = GatedClient('unsubscribe')
        ev = Events(client, 1)
        first = ev.on('Click', lambda *args: None)

        remover = threading.Thread(target=lambda: ev.off(first))
        remover.start()
        self.assertIsNotNone(client.wait_until_inside(), 'the remover must reach its wire call')

        adder = threading.Thread(target=lambda: ev.on('Click', lambda *args: None))
        adder.start()
        # Under the fix the adder cannot even begin to decide until the
        # remover is done with the wire, so this join is expected to time out.
        # It is not asserted either way -- what matters is the state both
        # orders end in.
        adder.join(1)
        client.let_it_through()
        remover.join(5)
        adder.join(5)
        self.assertFalse(remover.is_alive(), 'the remover must finish')
        self.assertFalse(adder.is_alive(), 'the adder must finish')

        self.assertEqual(ev._registered_names(), ['Click'], 'a callback is registered')
        self.assertEqual(
            client.calls[-1][0], 'subscribe',
            'so the last thing the bridge heard must be a subscribe -- an unsubscribe here is a '
            'registered callback whose event can never arrive')

    # The mirror: the bridge left advised with nothing to deliver to. A leaked
    # Advise, and every event it goes on raising minting handles nobody
    # releases.
    def test_an_unsubscribe_must_not_land_after_the_subscribe_that_follows_it(self):
        client = GatedClient('subscribe')
        ev = Events(client, 1)

        adder = threading.Thread(target=lambda: ev.on('Click', lambda *args: None))
        adder.start()
        self.assertIsNotNone(client.wait_until_inside(), 'the adder must reach its wire call')

        remover = threading.Thread(target=lambda: ev.off('Click'))
        remover.start()
        remover.join(1)
        client.let_it_through()
        adder.join(5)
        remover.join(5)
        self.assertFalse(adder.is_alive(), 'the adder must finish')
        self.assertFalse(remover.is_alive(), 'the remover must finish')

        self.assertEqual(ev._registered_names(), [], 'no callback is registered')
        self.assertEqual(
            client.calls[-1][0], 'unsubscribe',
            'so the last thing the bridge heard must be an unsubscribe -- a subscribe here leaves '
            'the object advised with nobody to deliver to')

    # The bulk form, for a caller who does not want to remember the names.
    def test_close_takes_every_subscription_and_the_thread_down(self):
        ev, client = self.events_for()
        ev.on('Click', lambda *args: None)
        ev.on('Other', lambda *args: None)
        ev.close()

        self.assertEqual(client.count('unsubscribe'), 2)
        self.assertEqual(ev._registered_names(), [])
        self.assertEqual(client.sinks, [])
        self.assertTrue(ev._dispatcher_stopped(5))

    # off removes the registration it was handed and no other. The same
    # function registered twice is two registrations, and value equality
    # cannot tell them apart: on a value-equal Subscription, off(first) would
    # remove both, send the unsubscribe, and the second would never fire again
    # with nothing said about it.
    def test_off_removes_only_the_registration_it_was_given(self):
        ev, client = self.events_for()
        seen = queue.Queue()

        def same_callback(*args):
            seen.put('ran')

        first = ev.on('Click', same_callback)
        second = ev.on('Click', same_callback)
        self.assertIsNot(first, second, 'two registrations, whatever they hold')

        ev.off(first)
        self.assertEqual(client.count('unsubscribe'), 0,
                         'the second registration is still there, so the subscription must stay')
        client.deliver(frame('Click', 1))
        self.assertEqual(waited(seen), 'ran', 'the registration that was not named must still fire')

    # Assert it fired BEFORE, then stopped. "Nothing arrived" on its own is
    # what a test that never worked also says. Two events, in order, on the
    # one dispatcher: if the removed callback still ran, its 'click' reaches
    # the queue first and the assertion says so.
    def test_off_stops_delivery(self):
        ev, client = self.events_for()
        seen = queue.Queue()
        click = ev.on('Click', lambda *args: seen.put('click'))
        # A second event keeps the object armed, so what this test measures is
        # off, not the teardown that follows the last callback.
        ev.on('Other', lambda *args: seen.put('other'))

        client.deliver(frame('Click', 1))
        self.assertEqual(waited(seen), 'click',
                         'it must fire before the off, or what follows proves nothing')

        ev.off(click)
        client.deliver(frame('Click', 2))
        client.deliver(frame('Other', 3))
        self.assertEqual(waited(seen), 'other', 'the removed callback must not have fired')

    # One connection's frames reach every Events on it, so names nobody
    # registered for arrive here all the time. A table that grew a permanent
    # key for each of them is a read that mutates.
    def test_an_event_nobody_registered_for_leaves_no_trace(self):
        ev, client = self.events_for()
        ev.on('Click', lambda *args: None)
        client.deliver(frame('Unwanted', 5))
        ev._drain()

        self.assertEqual(ev._registered_names(), ['Click'])

    # args: null says the bridge minted NOTHING for this event -- it is sent
    # as null rather than left out precisely so the client can tell that from
    # "zero arguments". release_event is a synchronous round trip from the
    # dispatcher, so sending one anyway caps the event rate at one RTT in
    # exactly the high-frequency case args=False exists for.
    def test_an_event_that_minted_nothing_is_not_released(self):
        ev, client = self.events_for()
        ran = queue.Queue()
        ev.on('Click', lambda *args: ran.put(args), args=False)
        client.deliver(frame('Click', 9))
        self.assertEqual(waited(ran), (),
                         'with the flag off the callback is called with no arguments at all')
        ev._drain()

        self.assertEqual(client.count('release_event'), 0, 'there is nothing to give back')

    # The decode this exercises -- the one that turns an event argument into
    # something a callback can use -- is the same one an invoke's result goes
    # through. A second, private copy of that walk is how those two answers
    # drift apart.
    def test_event_arguments_arrive_decoded_like_any_other_result(self):
        ev, client = self.events_for()
        got = queue.Queue()
        ev.on('Click', lambda *args: got.put(args))

        client.deliver({'event': 'Click', 'handle': 1, 'seq': 5, 'args': [
            {'$ole_ref': 77},
            {'$type': 'time', 'iso8601': '2026-08-31T09:30:45'},
            'plain',
            42,
        ]})
        args = waited(got)

        self.assertIsNotNone(args, 'the callback must have been called')
        self.assertIsInstance(args[0], Proxy, 'an object argument must arrive callable')
        self.assertEqual(args[0].ole_handle, 77)
        self.assertIsInstance(args[1], datetime.datetime,
                              'and a date as a datetime, exactly as a call result would')
        self.assertEqual(args[1].hour, 9)
        self.assertEqual(args[2], 'plain')
        self.assertEqual(args[3], 42)

    def test_callbacks_run_in_registration_order(self):
        ev, client = self.events_for()
        order = queue.Queue()
        ev.on('Click', lambda *args: order.put('first'))
        ev.on('Click', lambda *args: order.put('second'))
        client.deliver(frame('Click', 1))
        self.assertEqual(waited(order), 'first')
        self.assertEqual(waited(order), 'second')

    def test_a_raising_callback_reaches_on_error_and_delivery_continues(self):
        ev, client = self.events_for()
        errors = queue.Queue()
        done = queue.Queue()
        ev.on_error(lambda exc, _frame: errors.put(exc))

        def boom(*args):
            raise RuntimeError('boom')

        ev.on('Click', boom)
        ev.on('Click', lambda *args: done.put('ran'))

        client.deliver(frame('Click', 1))
        error = waited(errors)
        self.assertIsNotNone(error, 'on_error must be told')
        self.assertEqual(str(error), 'boom')
        self.assertEqual(waited(done), 'ran', 'a later callback must still run')

        client.deliver(frame('Click', 2))
        # boom runs before done (registration order), so this event's error
        # is already on the queue by the time done.put runs. Drained here, or
        # it would still be sitting there when the on_error(None) case below
        # checks the queue is empty -- and be mistaken for a handler that was
        # supposed to have been taken off.
        self.assertEqual(str(waited(errors)), 'boom', 'the still-installed handler is told again')
        self.assertEqual(waited(done), 'ran', 'and the next event must still be delivered')

        # on_error(None) restores the default: Python has no empty block to
        # pass, so None is the explicit form -- and it returns the Events, so
        # it chains, exactly as Ruby's does.
        self.assertIs(ev.on_error(None), ev)
        stderr = io.StringIO()
        with contextlib.redirect_stderr(stderr):
            client.deliver(frame('Click', 3))
            self.assertEqual(waited(done), 'ran')
            ev._drain()
        self.assertIsNone(waited(errors, 0.5),
                          'the handler that was taken off must not be called again')
        self.assertIn('callback raised RuntimeError: boom', stderr.getvalue(),
                      'without a handler the default reports it on stderr')

    # The arguments are callback-scoped, so the release has to happen even
    # when the callback blew up -- otherwise one bad callback leaks two COM
    # objects per event, forever.
    def test_event_handles_are_released_even_when_the_callback_raises(self):
        ev, client = self.events_for()
        ev.on_error(lambda exc, frame: None)

        def boom(*args):
            raise RuntimeError('boom')

        ev.on('Click', boom)
        client.deliver({'event': 'Click', 'handle': 1, 'seq': 77, 'args': []})
        ev._drain()
        self.assertIn(('release_event', {'seq': 77}), client.calls)

    # A dead dispatcher is permanent and silent, which is why the except
    # around a callback is BaseException and not the customary Exception.
    # Measured on the Ruby code that used the narrow one: a callback raising
    # outside the usual hierarchy ended the thread, and the 10 events that
    # followed were queued and never delivered -- callbacks still registered,
    # bridge still advised, 1 of 11 release_events sent, on_error never told.
    def test_a_callback_raising_past_standard_error_does_not_stop_the_dispatcher(self):
        ev, client = self.events_for()
        errors = queue.Queue()
        ran = queue.Queue()
        ev.on_error(lambda exc, _frame: errors.put(exc))

        def past_the_usual_except(*args):
            raise KeyboardInterrupt('past the usual except')

        ev.on('Click', past_the_usual_except)
        ev.on('Click', lambda *args: ran.put('later_callback'))

        client.deliver(frame('Click', 78, args=[]))
        error = waited(errors)
        self.assertIsNotNone(error, 'on_error must be told')
        self.assertIsInstance(error, KeyboardInterrupt)
        self.assertEqual(waited(ran), 'later_callback', 'the callbacks after it must still run')

        client.deliver(frame('Click', 79, args=[]))
        self.assertEqual(waited(ran), 'later_callback',
                         'and every later event must still be delivered')

        ev._drain()
        self.assertIn(('release_event', {'seq': 78}), client.calls)
        self.assertIn(('release_event', {'seq': 79}), client.calls)
        self.assertFalse(ev._dispatcher_stopped(0.1), 'the dispatcher thread must still be there')

    # The other way to lose the thread: a frame whose args is neither a list
    # nor null is something the bridge cannot have sent, and it is refused
    # before any callback is reached. In Python the check has to be written --
    # a str is iterable, so without it the callback would be handed one
    # argument per character. The handles the frame names were minted all the
    # same.
    def test_a_malformed_frame_reaches_on_error_and_still_releases_its_handles(self):
        ev, client = self.events_for()
        errors = queue.Queue()
        ran = queue.Queue()
        ev.on_error(lambda exc, _frame: errors.put(exc))
        ev.on('Click', lambda *args: ran.put('ran'))

        client.deliver({'event': 'Click', 'handle': 1, 'seq': 80, 'args': 'not a list'})
        error = waited(errors)
        self.assertIsInstance(error, TypeError)
        self.assertIsNone(waited(ran, 0.5), 'no callback may be called with a frame like that')
        ev._drain()
        self.assertIn(('release_event', {'seq': 80}), client.calls,
                      'the handles the bridge minted for it must still go back')

        client.deliver(frame('Click', 81, args=[]))
        self.assertEqual(waited(ran), 'ran', 'and the next event must still be delivered')

    # The third way it was lost: the reporting machinery's own except was the
    # narrow one, so an on_error raising past it killed the thread from inside
    # the code that exists to report exactly this.
    def test_an_on_error_that_itself_raises_does_not_stop_the_dispatcher(self):
        ev, client = self.events_for()
        ran = queue.Queue()

        def broken_handler(_exc, _frame):
            raise KeyboardInterrupt('the handler is broken too')

        def boom(*args):
            raise RuntimeError('boom')

        ev.on_error(broken_handler)
        ev.on('Click', boom)
        ev.on('Click', lambda *args: ran.put('later_callback'))

        with contextlib.redirect_stderr(io.StringIO()):
            client.deliver(frame('Click', 1))
            self.assertEqual(waited(ran), 'later_callback', 'the callbacks after it must still run')
            client.deliver(frame('Click', 2))
            self.assertEqual(waited(ran), 'later_callback',
                             'and every later event must still be delivered')

    # A callback runs on the dispatcher, and the dispatcher is what off takes
    # down when the last callback goes. Registering and removing from inside
    # one must therefore work, and must not deadlock against the lock the
    # callback's own delivery took on the way in.
    def test_a_callback_can_register_and_remove_from_inside_the_dispatcher(self):
        ev, client = self.events_for()
        survived = queue.Queue()
        seen = queue.Queue()
        holder = {}

        def callback(*args):
            # This order on purpose: the off empties the registry, which takes
            # the thread this callback is running on down with it, and the on
            # puts it straight back. One dispatcher must come out of that, not
            # two and not none.
            ev.off(holder['sub'])
            ev.on('Other', lambda *rest: seen.put('other'))
            survived.put('survived')

        holder['sub'] = ev.on('Click', callback)

        client.deliver(frame('Click', 1))
        self.assertEqual(waited(survived), 'survived',
                         'a callback calling off/on must not deadlock')
        client.deliver(frame('Other', 2))
        self.assertEqual(waited(seen), 'other', 'and the object must still be delivering')

    # An Events that has had its last callback removed used to keep a parked
    # dispatcher thread and an entry the reader walks for every frame on the
    # connection -- measured on Ruby: 50 proxies that registered one callback
    # and removed it left 51 live threads and 50 sink entries, until close.
    # Nothing is registered, so nothing is derived.
    def test_the_last_callback_removed_takes_the_dispatcher_and_the_sink_with_it(self):
        client = FakeClient()
        ev = Events(client, 1)
        self.assertEqual(client.sinks, [], 'an ole_events nobody has registered on costs nothing')

        seen = queue.Queue()
        sub = ev.on('Click', lambda *args: seen.put('ran'))
        self.assertEqual(len(client.sinks), 1)
        client.deliver(frame('Click', 1))
        self.assertEqual(waited(seen), 'ran',
                         'it must deliver before the off, or what follows proves nothing')

        ev.off(sub)
        self.assertTrue(ev._dispatcher_stopped(5),
                        'the dispatcher must not stay parked with nothing to deliver')
        self.assertEqual(client.sinks, [], 'and the sink must come off the connection')

        ev.on('Click', lambda *args: seen.put('again'))
        client.deliver(frame('Click', 2))
        self.assertEqual(waited(seen), 'again',
                         'registering again must work exactly as the first time did')

    # One connection can carry several objects with events. A registration
    # that REPLACED the previous consumer would switch the earlier object's
    # events off in silence.
    def test_two_event_objects_on_one_connection_both_receive(self):
        client = FakeClient()
        a = Events(client, 1)
        b = Events(client, 2)
        seen = queue.Queue()
        a.on('Click', lambda *args: seen.put('a'))
        b.on('Click', lambda *args: seen.put('b'))

        client.deliver(frame('Click', 1, handle=2))
        self.assertEqual(waited(seen), 'b')
        client.deliver(frame('Click', 2, handle=1))
        self.assertEqual(waited(seen), 'a')

    # The promise the README makes -- "one dispatcher thread per connection,
    # in arrival order, one at a time" -- and the whole value of it: a caller
    # who shares a dict between an Application callback and a Workbook
    # callback needs no lock, because the two can never be inside their
    # callbacks at once. Measured on Ruby with a thread per Events object:
    # both callbacks were in their blocks simultaneously, on 2 distinct
    # threads, with every other assertion in the file still passing.
    def test_two_event_objects_on_one_connection_share_one_thread_and_never_overlap(self):
        client = FakeClient()
        a = Events(client, 1)
        b = Events(client, 2)
        entered = queue.Queue()
        gate = queue.Queue()
        threads = queue.Queue()

        def callback_a(*args):
            threads.put(threading.current_thread())
            entered.put('a')
            gate.get()

        def callback_b(*args):
            threads.put(threading.current_thread())
            entered.put('b')

        a.on('Click', callback_a)
        b.on('Click', callback_b)

        client.deliver(frame('Click', 1, handle=1))
        self.assertEqual(waited(entered), 'a', 'the first callback must be inside its callback')

        client.deliver(frame('Click', 2, handle=2))
        self.assertIsNone(
            waited(entered, 0.5),
            "the second object's callback must not run while the first is held: one at a time")

        gate.put('go')
        self.assertEqual(waited(entered), 'b', 'and it must run as soon as the first returns')
        first_thread = waited(threads)
        second_thread = waited(threads)
        self.assertIsNotNone(first_thread)
        self.assertIs(first_thread, second_thread,
                      "both callbacks must have run on the connection's ONE dispatcher thread -- a "
                      "thread per object is a data race in every callback that shares state with "
                      "another object's")

    # off means off, even for what is already in flight. A frame minted before
    # the unsubscribe reached the bridge names handles the bridge is holding,
    # so it is released rather than delivered -- dropping it silently leaks
    # two COM objects per event for the life of the connection.
    def test_a_frame_for_a_handle_with_no_target_is_released_not_delivered(self):
        client = FakeClient()
        ev = Events(client, 1)
        ran = queue.Queue()
        ev.on('Click', lambda *args: ran.put('ran'))

        client.deliver(frame('Click', 42, args=[], handle=2))
        ev._drain()

        self.assertTrue(ran.empty(), 'a frame names one object, and no other object may see it')
        self.assertIn(('release_event', {'seq': 42}), client.calls,
                      'the handles the bridge minted for it must still go back')

    # The connection's dispatcher belongs to the connection, not to whichever
    # object armed it first. One object going quiet must leave the others
    # exactly as they were -- same thread, same queue, no restart.
    def test_detaching_one_object_keeps_delivering_to_the_other_on_the_same_thread(self):
        client = FakeClient()
        a = Events(client, 1)
        b = Events(client, 2)
        seen = queue.Queue()
        threads = queue.Queue()
        a.on('Click', lambda *args: seen.put('a'))

        def callback_b(*args):
            threads.put(threading.current_thread())
            seen.put('b')

        b.on('Click', callback_b)

        client.deliver(frame('Click', 1, handle=2))
        self.assertEqual(waited(seen), 'b',
                         'it must deliver before the off, or what follows proves nothing')
        before = waited(threads)

        a.off('Click')
        client.deliver(frame('Click', 2, handle=2))
        self.assertEqual(waited(seen), 'b',
                         "one object's last callback must not stop the connection's other "
                         'deliveries')
        self.assertIs(before, waited(threads), 'and it must still be the same dispatcher thread')
        self.assertFalse(b._dispatcher_stopped(0.1), 'which is therefore still running')

    # A frame queued after an Events re-armed the connection during the
    # idle hand-off must reach it, on the same dispatcher thread: the thread
    # that read the idle marker gives back what was left over (the round trip
    # this test holds open) and then resumes rather than exiting.
    def test_a_frame_arriving_during_the_idle_hand_off_reaches_the_new_target(self):
        client = OnceReleaseGatedClient()
        a = Events(client, 1)
        in_callback = queue.Queue()
        hold = queue.Queue()

        def slow(*args):
            in_callback.put('in')
            hold.get()

        a.on('Click', slow)
        thread = client.dispatcher._dispatcher_thread()
        self.assertIsNotNone(thread, 'the first registration must have started a dispatcher')

        # Held inside a callback, which is the only way to build a queue up
        # behind the dispatcher.
        client.deliver(frame('Click', 1))
        self.assertEqual(waited(in_callback), 'in')

        # off, on, a frame, off: the frame ends up behind an idle marker, and
        # giving it back is the round trip that holds the thread.
        a.off('Click')
        a.on('Click', lambda *args: None)
        client.deliver(frame('Click', 99, args=[]))
        a.off('Click')

        hold.put('go')
        self.assertEqual(
            client.wait_until_inside(), 99,
            'the dispatcher must be inside the release of what it had left over')

        # Re-arm while it is held there. No second thread may appear.
        b = Events(client, 2)
        got = queue.Queue()
        b.on('Click', lambda *args: got.put('b'))
        self.assertIs(client.dispatcher._dispatcher_thread(), thread,
                      'a re-arm during the hand-off must not start a second dispatcher')

        # Queued after the re-arm: delivered once the thread resumes.
        client.deliver(frame('Click', 77, handle=2))
        client.let_it_through()

        self.assertEqual(
            waited(got), 'b',
            'a frame for an object that attached during the idle hand-off must reach it')
        self.assertIs(client.dispatcher._dispatcher_thread(), thread,
                      'and it must have been delivered by the thread that resumed, not a new one')
        self.assertEqual(
            [params for method, params in client.calls if method == 'release_event'],
            [{'seq': 99}],
            'and what WAS left over from before the hand-off must still be given back')
        b.close()
        self.assertTrue(b._dispatcher_stopped(5))

    # The dispatcher's targets are keyed by handle, and two Events on one
    # handle would otherwise unseat each other: the second attach overwrites
    # the first's slot, and either one's last off deletes it -- taking the
    # other's routing and the connection's dispatcher thread with it.
    #
    # Proxy.ole_events memoises and bridge ids are unique per session, so
    # nothing shipped can reach this. But Events(client, handle) is public,
    # every test in this file builds them that way, and "removing the wrong
    # one silently stops a live consumer" is exactly what off_event refuses to
    # leave to convention.
    def test_a_second_events_on_the_same_handle_is_refused_rather_than_unseating_the_first(self):
        client = FakeClient()
        first = Events(client, 1)
        second = Events(client, 1)
        seen = queue.Queue()
        first.on('Click', lambda *args: seen.put('first'))

        with self.assertRaises(ValueError) as ctx:
            second.on('Click', lambda *args: seen.put('second'))
        self.assertIn('handle 1', str(ctx.exception),
                      'the refusal must name the handle it is about')

        self.assertEqual(len(client.sinks), 1,
                         'the refused object must leave nothing on the connection')
        client.deliver(frame('Click', 2))
        self.assertEqual(waited(seen), 'first',
                         'the object that was there first must still receive')
        self.assertFalse(first._dispatcher_stopped(0.1),
                         "and the connection's dispatcher must still be running")

    # A dispatcher blocked on an empty queue never exits, and every Events
    # leaks a thread for the life of the process.
    def test_the_dispatcher_thread_stops_when_the_connection_ends(self):
        client = FakeClient()
        ev = Events(client, 1)
        ev.on('Click', lambda *args: None)
        client.deliver(None)
        self.assertTrue(ev._dispatcher_stopped(2),
                        'the dispatcher must finish after the stream ends')

    # Registering and keeping no reference is the ordinary shape --
    # xl.ole_events.on('Click', cb) leaves the caller holding nothing but the
    # Proxy, and a caller who drops that holds nothing at all. The Client
    # holds the Dispatcher, and the Dispatcher's target table holds the
    # Events: that is the whole mechanism, and the opposite direction of the
    # collectability test below. Break it -- hold the targets weakly -- and
    # the Events is collected out from under a live connection: callbacks stop
    # firing, the bridge stays advised, and every event it goes on sending
    # leaks its argument handles because nobody is left to release them.
    def register_and_forget(self, client, seen):
        # In a method of its own so the Events is genuinely unreferenced -- a
        # local in the test body would keep it alive and the test would pass
        # whatever the dispatcher held.
        Events(client, 1).on('Click', lambda *args: seen.put('ran'))

    def test_a_registered_callback_survives_a_garbage_collection(self):
        client = FakeClient()
        seen = queue.Queue()
        self.register_and_forget(client, seen)
        gc.collect()

        client.deliver(frame('Click', 1))
        self.assertEqual(waited(seen), 'ran',
                         'the callback must outlive the reference the caller dropped')

    # A Client whose socket is never closed is a leak this project has found
    # three times, at three layers. A dispatcher thread started from a bound
    # method or a closure over self walks it straight back in: the running
    # thread holds its target, which roots the Dispatcher, which holds the
    # Client, whose __del__ therefore never runs and whose peer never sees
    # EOF.
    def test_a_client_with_events_on_it_is_still_collectable(self):
        server, client_side = socket.socketpair()
        saw_eof = []

        def responder():
            # Answers whatever `on` sends, so subscribing can complete. It
            # holds the server end only -- nothing that could pin the client
            # under test.
            reader = server.makefile('rb')
            try:
                while True:
                    line = reader.readline()
                    if not line:
                        saw_eof.append(True)
                        return
                    request = json.loads(line)
                    server.sendall(
                        (json.dumps({'id': request['id'], 'result': True}) + '\n').encode())
            except (OSError, ValueError):
                return

        responder_thread = threading.Thread(target=responder, daemon=True)
        responder_thread.start()
        carried = {}

        def build():
            # Built inside a function, so the test's own frame never holds a
            # reference to the Client or the Events.
            client = Client(client_side)
            ev = Events(client, 1)
            ev.on('Click', lambda *args: None)
            # The Thread object, not a bound method of the dispatcher: a bound
            # method would pin the very object this test is about. A Thread
            # holds only what was handed to it -- a weakref and the queue.
            carried['dispatcher'] = client.dispatcher._dispatcher_thread()
            return weakref.ref(client)

        try:
            weak = build()
            # The Client/Dispatcher/Events ring is a cycle, so refcounting
            # alone never frees it and this is what does.
            gc.collect()

            self.assertIsNone(
                weak(),
                'a Client with an Events on it must still be collectible -- something is pinning it')
            responder_thread.join(5)
            self.assertFalse(
                responder_thread.is_alive(),
                'the peer must see EOF: __del__ only closes the socket if the Client was collected')
            self.assertEqual(saw_eof, [True])
            carried['dispatcher'].join(5)
            self.assertFalse(
                carried['dispatcher'].is_alive(),
                'a collected Dispatcher must not leave its thread parked on a queue nothing can '
                'push to')
        finally:
            server.close()

    # Closing the connection is the ordinary thing to do from a callback ("I
    # have seen what I was waiting for"), and it is a shutdown reaching in
    # from a thread the shutdown itself has to stop. It must return rather
    # than deadlock: on a real Client, over a real socket, with a real reader
    # thread to join.
    def test_a_callback_can_close_the_client_that_delivered_to_it(self):
        with FakeBridge(handler=lambda method, params: True) as bridge:
            client = Client(bridge.sock)
            ev = Events(client, 1)
            closed = queue.Queue()

            def callback(*args):
                try:
                    client.close()
                    closed.put('returned')
                except Exception as exc:
                    closed.put(exc)

            ev.on('Click', callback)
            bridge.push({'event': 'Click', 'handle': 1, 'seq': 1, 'args': None})

            result = waited(closed)
            self.assertEqual(result, 'returned',
                             f"close from inside a callback must return (got {result!r})")


if __name__ == '__main__':
    unittest.main()
