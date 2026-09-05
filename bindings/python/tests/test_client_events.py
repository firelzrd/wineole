import contextlib
import gc
import io
import json
import os
import queue
import sys
import threading
import unittest

sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from fake_bridge import FakeBridge
from wineole.client import Client
from wineole.errors import ProtocolError


def waited(box, seconds=5):
    """A bounded get. Under every bug these tests name -- a frame that never
    reaches a consumer, a waiter that is never woken -- nothing is ever put,
    so an unbounded get would hang the suite instead of failing it. Returns
    None on a miss, which fails the assertion that follows in seconds."""
    try:
        return box.get(timeout=seconds)
    except queue.Empty:
        return None


class AnswerThenBreak:
    """A socket stand-in that answers one request and then ends the stream by
    raising, with no wait in between: the window in which the reader has
    filled a waiter and woken its caller, but the caller has not yet
    re-acquired the mailbox lock to delete itself from the table."""

    def __init__(self):
        self._requests = queue.Queue()
        self._answered = False

    def sendall(self, data):
        self._requests.put(data)

    def makefile(self, mode='rb'):
        return self

    def readline(self):
        if self._answered:
            raise OSError('the stream ends here, with nothing to wait for')
        raw = self._requests.get()
        if raw is None:
            return b''
        self._answered = True
        return (json.dumps({'id': json.loads(raw)['id'], 'result': 'answered'}) + '\n').encode()

    def shutdown(self, how):
        pass

    def close(self):
        self._requests.put(None)


class ClientEventsTest(unittest.TestCase):
    @contextlib.contextmanager
    def fake_bridge(self, handler=None):
        with FakeBridge(handler=handler) as bridge:
            client = Client(bridge.sock)
            try:
                yield client, bridge
            finally:
                # close() is bounded internally (it joins the reader for at
                # most 2s), so a client whose close cannot complete fails the
                # test rather than hanging the suite.
                try:
                    client.close()
                except Exception:
                    pass

    def test_an_event_frame_does_not_break_a_pending_response(self):
        with self.fake_bridge() as (client, bridge):
            got = []
            caller = threading.Thread(target=lambda: got.append(client.call('invoke', {'handle': 1})))
            caller.start()

            request = bridge.take_request()
            self.assertIsNotNone(request, 'the request must reach the bridge')
            # The event arrives FIRST, before the response. The old client
            # read the next line as its own response and raised "id
            # mismatch"; that is the regression this pins.
            bridge.push({'event': 'SheetChange', 'handle': 1, 'seq': 5, 'args': None})
            bridge.reply(request, 42)

            caller.join(5)
            self.assertFalse(caller.is_alive(), 'the caller must be answered')
            self.assertEqual(got, [42])

    def test_events_reach_the_registered_consumer(self):
        with self.fake_bridge() as (client, bridge):
            seen = queue.Queue()

            def sink(frame):
                seen.put([frame])

            client.on_event(sink)
            bridge.push({'event': 'Click', 'handle': 3, 'seq': 9, 'args': None})

            box = waited(seen)
            self.assertIsNotNone(box, 'the consumer must be handed the frame')
            # The whole frame, not a field at a time: what this client owes a
            # consumer is the parsed frame exactly as it arrived, with nothing
            # dropped, renamed or filled in.
            self.assertEqual(box[0], {'event': 'Click', 'handle': 3, 'seq': 9, 'args': None},
                             'the frame must reach the consumer exactly as it arrived')

    def test_on_event_appends_rather_than_replacing_the_consumer(self):
        # One connection can carry several objects with events on them (an
        # Application and a Workbook, say), each with its own consumer. A
        # registration that REPLACED the previous one would silently switch
        # the earlier objects' events off.
        with self.fake_bridge() as (client, bridge):
            first = queue.Queue()
            second = queue.Queue()
            client.on_event(lambda frame: first.put([frame]))
            client.on_event(lambda frame: second.put([frame]))

            bridge.push({'event': 'Click', 'handle': 3, 'seq': 9, 'args': None})

            box = waited(first)
            self.assertIsNotNone(box, 'the consumer registered first must still receive events')
            self.assertEqual(box[0]['event'], 'Click')
            box = waited(second)
            self.assertIsNotNone(box, 'the consumer registered second must receive events too')
            self.assertEqual(box[0]['event'], 'Click')

    def test_off_event_stops_one_consumer_and_leaves_the_others(self):
        # A consumer registered for the life of the connection cannot be
        # dismantled. Events takes its sink off when its last callback goes,
        # and without this the reader would walk an entry for every object
        # that ever had a callback on it.
        with self.fake_bridge() as (client, bridge):
            dropped = queue.Queue()
            kept = queue.Queue()

            # Registered FIRST, so it is called first while it is registered:
            # if it were still there after the off, its frame would be in the
            # queue before the one this test waits for.
            def going(frame):
                dropped.put([frame])

            client.on_event(going)
            client.on_event(lambda frame: kept.put([frame]))

            bridge.push({'event': 'Click', 'handle': 3, 'seq': 9, 'args': None})
            box = waited(dropped)
            self.assertIsNotNone(box, 'it must receive before the off, or what follows proves nothing')
            self.assertEqual(box[0]['event'], 'Click')
            self.assertIsNotNone(waited(kept))

            client.off_event(going)
            bridge.push({'event': 'Other', 'handle': 3, 'seq': 10, 'args': None})

            box = waited(kept)
            self.assertIsNotNone(box, 'the consumer that stayed must go on receiving')
            self.assertEqual(box[0]['event'], 'Other')
            self.assertTrue(dropped.empty(), 'the removed consumer must receive nothing further')
            self.assertEqual(len(client._event_sinks), 1,
                             'and it must not be left holding the connection')

    def test_two_threads_can_have_requests_in_flight_at_once(self):
        # The old client held a lock across the whole round trip, so a call
        # made from inside an event callback could not even be sent until the
        # outer call returned.
        with self.fake_bridge() as (client, bridge):
            results = {}
            a = threading.Thread(target=lambda: results.setdefault('a', client.call('invoke', {'n': 1})))
            a.start()
            first = bridge.take_request()
            self.assertIsNotNone(first, 'the first request must reach the wire')

            b = threading.Thread(target=lambda: results.setdefault('b', client.call('invoke', {'n': 2})))
            b.start()
            second = bridge.take_request()
            self.assertIsNotNone(
                second, 'the second request must reach the wire while the first is unanswered')

            # Answer them out of order, to prove the routing is by id and not
            # by arrival order.
            bridge.reply(second, 'second')
            bridge.reply(first, 'first')

            a.join(5)
            b.join(5)
            self.assertFalse(a.is_alive(), 'the first caller must be answered')
            self.assertFalse(b.is_alive(), 'the second caller must be answered')
            self.assertEqual(results, {'a': 'first', 'b': 'second'})

    def test_a_closed_connection_wakes_every_waiter(self):
        with self.fake_bridge() as (client, bridge):
            outcomes = []

            def call_it():
                try:
                    client.call('invoke', {})
                except Exception as exc:
                    outcomes.append(exc)

            waiters = [threading.Thread(target=call_it) for _ in range(3)]
            for w in waiters:
                w.start()
            for _ in range(3):
                self.assertIsNotNone(bridge.take_request(), 'every request must reach the wire')

            bridge.close()

            for w in waiters:
                w.join(5)
                self.assertFalse(w.is_alive(), 'every waiter must be woken on EOF, or it waits forever')
            self.assertEqual(len(outcomes), 3)
            for exc in outcomes:
                self.assertIsInstance(exc, ProtocolError)

    def test_a_raising_event_consumer_does_not_take_the_connection_down(self):
        # Every consumer on a connection shares the one reader thread, so an
        # exception raised out of one of them would otherwise end the read
        # loop: the other consumers would go silent and every later call would
        # fail with "connection closed".
        with contextlib.redirect_stderr(io.StringIO()):
            with self.fake_bridge() as (client, bridge):
                seen = queue.Queue()

                def broken(_frame):
                    raise RuntimeError('this consumer is broken')

                client.on_event(broken)
                client.on_event(lambda frame: seen.put([frame]))

                bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})
                box = waited(seen)
                self.assertIsNotNone(
                    box, 'a consumer registered after a broken one must still see events')
                self.assertEqual(box[0]['event'], 'Click')

                # ...and the connection must still carry a round trip.
                answer = []
                pending = threading.Thread(target=lambda: answer.append(client.call('invoke', {})))
                pending.start()
                request = bridge.take_request()
                self.assertIsNotNone(request)
                bridge.reply(request, 'alive')
                pending.join(5)
                self.assertFalse(pending.is_alive(), 'the connection must still answer a call')
                self.assertEqual(answer, ['alive'])

    def test_a_closed_connection_tells_every_event_consumer_the_stream_ended(self):
        # Events parks a dispatcher thread on a queue that only the reader
        # fills. Without an end-of-stream hand-off that thread blocks on an
        # empty queue for the life of the process.
        with self.fake_bridge() as (client, bridge):
            first = queue.Queue()
            second = queue.Queue()
            client.on_event(lambda frame: first.put([frame]))
            client.on_event(lambda frame: second.put([frame]))

            bridge.close()

            box = waited(first)
            self.assertIsNotNone(box, 'the first consumer must be told')
            self.assertIsNone(box[0], 'the first consumer must be handed None at EOF')
            box = waited(second)
            self.assertIsNotNone(box, 'the second consumer must be told')
            self.assertIsNone(box[0], 'the second consumer must be handed None at EOF')

    def test_a_consumer_registered_after_the_stream_ended_is_told_so_at_once(self):
        # The realistic path: the bridge dies, and only then does user code
        # attach a handler to a proxy it already had.
        with self.fake_bridge() as (client, bridge):
            ended = queue.Queue()
            client.on_event(lambda frame: ended.put([frame]))
            bridge.close()
            box = waited(ended)
            self.assertIsNotNone(box)
            self.assertIsNone(box[0], 'the stream must have ended')

            late = queue.Queue()
            client.on_event(lambda frame: late.put([frame]))

            box = waited(late)
            self.assertIsNotNone(
                box, 'a consumer registered after the stream ended must be handed None immediately')
            self.assertIsNone(box[0])

    def test_a_call_issued_from_inside_an_event_consumer_reaches_the_wire(self):
        # The deadlock this whole design exists to prevent. Holding the
        # mailbox lock across the dispatch leaves every other test in this
        # file passing, and yet a call issued from inside a consumer cannot
        # even reach the wire: the reader holds the lock that guards the
        # socket write, on the very thread the consumer runs on.
        with self.fake_bridge() as (client, bridge):
            def sink(frame):
                if frame is None:
                    return
                try:
                    # The reader thread is inside this block, so nobody is
                    # left to route the response back and this call can never
                    # complete -- it returns when the fixture closes the
                    # client, which fails every pending waiter. That is the
                    # documented contract (hand off, do not block); what is
                    # under test is only that the request got out.
                    client.call('invoke', {'from': 'a consumer'})
                except Exception:
                    pass

            client.on_event(sink)
            bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})

            request = bridge.take_request()
            self.assertIsNotNone(request, 'a call made from inside an event consumer must reach the wire')
            self.assertEqual(request['method'], 'invoke')

    def test_only_frames_without_an_id_reach_the_event_consumers(self):
        # Frames WITH an id are answers to a caller and nobody else's
        # business. Dispatching them to the sinks as well leaves every other
        # test here passing, while a consumer that reads frame['handle'] would
        # raise once per RPC -- silently, since a raising sink is only warned
        # about.
        with self.fake_bridge() as (client, bridge):
            seen = queue.Queue()
            client.on_event(lambda frame: seen.put([frame]))

            answer = []
            pending = threading.Thread(target=lambda: answer.append(client.call('invoke', {})))
            pending.start()
            request = bridge.take_request()
            self.assertIsNotNone(request)
            # The response FIRST, then the event: whatever reaches the
            # consumer first is what the reader considers dispatchable.
            bridge.reply(request, 'answered')
            bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})

            pending.join(5)
            self.assertFalse(pending.is_alive(), 'the caller must be answered')
            self.assertEqual(answer, ['answered'])

            box = waited(seen)
            self.assertIsNotNone(box)
            self.assertEqual(box[0]['event'], 'Click',
                             'a response frame must go to its waiter only, never to the consumers')
            self.assertTrue(seen.empty(), 'nothing but the event frame may reach a consumer')

    def test_a_valid_json_line_that_is_not_an_object_does_not_kill_the_reader(self):
        # null, 123 and [] are all valid JSON and none of them is a frame. An
        # unparseable line was already skipped; a parseable non-object used to
        # raise out of the frame lookup, killing the reader and with it every
        # consumer and every later call on the connection.
        with self.fake_bridge() as (client, bridge):
            seen = queue.Queue()
            client.on_event(lambda frame: seen.put([frame]))

            for raw in ('null', '123', '[]', '"a string"', 'not json at all'):
                bridge.push_line(raw)
            bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})

            box = waited(seen)
            self.assertIsNotNone(box, 'a bad line killed the reader: nothing reached the consumer')
            self.assertIsNotNone(box[0], 'the consumer was handed end-of-stream instead of the event')
            self.assertEqual(box[0]['event'], 'Click',
                             'a parseable non-object line must be skipped, not fatal')

            answer = []
            pending = threading.Thread(target=lambda: answer.append(client.call('invoke', {})))
            pending.start()
            request = bridge.take_request()
            self.assertIsNotNone(request, 'the reader must have survived the bad lines')
            bridge.reply(request, 'alive')
            pending.join(5)
            self.assertEqual(answer, ['alive'])

    def test_a_frame_whose_id_is_not_hashable_does_not_kill_the_reader(self):
        # A dict lookup on an unhashable key raises TypeError, not KeyError --
        # {}.get([1, 2]) blows up before it can even fail to find anything.
        # Uncaught, that kills the reader and with it every consumer and
        # every later call on the connection.
        with self.fake_bridge() as (client, bridge):
            seen = queue.Queue()
            client.on_event(lambda frame: seen.put([frame]))

            bridge.push_line(json.dumps({'id': [1, 2], 'result': 1}))
            bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})

            box = waited(seen)
            self.assertIsNotNone(box, 'an unhashable id killed the reader: nothing reached the consumer')
            self.assertIsNotNone(box[0], 'the consumer was handed end-of-stream instead of the event')
            self.assertEqual(box[0]['event'], 'Click',
                             'a frame with an unhashable id must be skipped, not fatal')

            answer = []
            pending = threading.Thread(target=lambda: answer.append(client.call('invoke', {})))
            pending.start()
            request = bridge.take_request()
            self.assertIsNotNone(request, 'the reader must have survived the unhashable id')
            bridge.reply(request, 'alive')
            pending.join(5)
            self.assertEqual(answer, ['alive'])

    def test_close_called_from_inside_an_event_consumer_does_not_raise(self):
        # A consumer runs on the reader thread, so close from inside one is
        # that thread joining itself -- which would raise out of the one
        # method whose job is to shut the connection down cleanly.
        with self.fake_bridge() as (client, bridge):
            outcome = queue.Queue()

            def sink(frame):
                if frame is None:
                    return
                try:
                    client.close()
                    outcome.put('closed')
                except Exception as exc:
                    outcome.put(exc)

            client.on_event(sink)
            bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})

            result = waited(outcome)
            self.assertEqual(result, 'closed',
                             f"close from inside a consumer must not raise (got {result!r})")

    def test_a_registered_consumer_survives_a_garbage_collection(self):
        # The reader reaches its sinks through weak references, so something
        # has to hold them: the Client does (Client.on_event). Drop that and a
        # single client.on_event(...) keeps working only until the next
        # collection, after which the connection silently stops delivering.
        with self.fake_bridge() as (client, bridge):
            seen = queue.Queue()
            client.on_event(lambda frame: seen.put([frame]))

            gc.collect()

            bridge.push({'event': 'Click', 'handle': 3, 'seq': 1, 'args': None})
            box = waited(seen)
            self.assertIsNotNone(box, 'the connection ended instead of delivering the event')
            self.assertIsNotNone(box[0])
            self.assertEqual(box[0]['event'], 'Click',
                             'a sink must live as long as the client it was registered on')

    def test_a_call_after_the_stream_ended_raises_rather_than_waiting_forever(self):
        # "The reader is dead but the socket is still open" -- the bridge
        # stopped answering, or half-closed. Without the closed check in
        # request the write succeeds, no reader is left to route an answer,
        # and the caller waits on its slot for the life of the process.
        with self.fake_bridge() as (client, bridge):
            ended = queue.Queue()
            client.on_event(lambda frame: ended.put([frame]))
            bridge.close_write()
            box = waited(ended)
            self.assertIsNotNone(box)
            self.assertIsNone(box[0], 'the reader must have finished')

            errors = []

            def call_it():
                try:
                    client.call('invoke', {})
                except Exception as exc:
                    errors.append(exc)

            caller = threading.Thread(target=call_it)
            caller.start()
            caller.join(5)
            self.assertFalse(caller.is_alive(), 'a call with no reader left must raise, not wait forever')
            self.assertEqual(len(errors), 1)
            self.assertIsInstance(errors[0], ProtocolError)
            self.assertEqual(str(errors[0]), 'connection closed')

    def test_the_end_of_the_stream_does_not_overwrite_a_delivered_response(self):
        # The sweep at EOF must not overwrite an answer that was already
        # delivered. Last-write-wins in the waiter turns a request the bridge
        # really answered into ProtocolError: connection closed.
        client = Client(AnswerThenBreak())
        outcome = {}

        def call_it():
            try:
                outcome['result'] = client.call('invoke', {})
            except Exception as exc:
                outcome['error'] = exc

        caller = threading.Thread(target=call_it)
        caller.start()
        caller.join(5)
        try:
            self.assertFalse(caller.is_alive(), 'the caller must not be left waiting')
            self.assertNotIn('error', outcome,
                             f"a request the bridge answered must not be reported as {outcome.get('error')!r}")
            self.assertEqual(outcome.get('result'), 'answered')
        finally:
            try:
                client.close()
            except Exception:
                pass


if __name__ == '__main__':
    unittest.main()
