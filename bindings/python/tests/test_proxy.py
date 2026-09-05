import copy
import datetime
import itertools
import os
import pickle
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.proxy import Proxy
from wineole.errors import NotSerializableError, StaleReferenceError
from wineole.events import Events

_fake_client_generation = itertools.count(1)


class FakeDispatcher:
    """Only what Proxy asks of a dispatcher: register a cleanup closure
    against a handle, and say afterwards what it was handed."""

    def __init__(self):
        self.registered_cleanups = []

    def register_cleanup(self, handle, fn):
        self.registered_cleanups.append((handle, fn))


class FakeClient:
    def __init__(self):
        self.calls = []
        # A fresh, distinct generation per instance -- mirroring
        # wineole/client.py's Client.generation (an explicit per-instance
        # counter, deliberately not id()). A hardcoded `self.generation = 1`
        # here would make every FakeClient indistinguishable from every
        # other, silently defeating any test that checks cross-client
        # behavior (two different clients must never compare equal).
        self.generation = next(_fake_client_generation)
        self.connect_or_create_created = True
        self.dispatcher = FakeDispatcher()
        self.await_cleanup_calls = []
        # What the bridge answers a `release` with. A dict carrying a
        # 'cleanup' key is how it says "a client closure has to run first".
        self.release_reply = None

    def await_cleanup(self, seq):
        self.await_cleanup_calls.append(seq)

    def call(self, method, params):
        self.calls.append((method, params))
        if method == 'create':
            return {'$ole_ref': 1}
        if method == 'connect':
            return {'$ole_ref': 1}
        if method == 'connect_or_create':
            return {'$ole_ref': 1, 'created': self.connect_or_create_created}
        if method == 'const_load':
            return {'xlUp': -4162, 'xlDown': -4121}
        if method == 'release':
            return self.release_reply
        if method == 'invoke':
            name = params['name']
            if name == 'Version':
                return 11.0
            if name == 'Worksheets':
                return {'$ole_ref': 2}
            if name == 'Timestamp':
                return {'$type': 'time', 'iso8601': '2026-08-26T04:00:00+00:00'}
            if name == 'BulkValue':
                return [
                    [1, {'$type': 'time', 'iso8601': '2026-08-31T09:30:00'}],
                    [None, 'text'],
                ]
            if name == 'BulkRefs':
                return [{'$ole_ref': 42}, {'$ole_ref': 43}]
            if name == 'NestedHash':
                return {'when': {'$type': 'time', 'iso8601': '2026-08-31T09:30:00'}}
            return None
        return None


class ProxyTest(unittest.TestCase):
    def test_attribute_access_returns_a_callable_not_an_immediate_value(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        member = proxy.Version  # bare attribute access -- must NOT call the client yet
        self.assertEqual(len(client.calls), 1)  # only the 'create' call so far
        result = member()  # calling it performs the RPC
        self.assertEqual(result, 11.0)
        self.assertEqual(len(client.calls), 2)

    def test_calling_a_member_decodes_an_ole_ref_into_a_new_proxy(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        sheets = proxy.Worksheets()

        self.assertIsInstance(sheets, Proxy)
        self.assertEqual(sheets.ole_handle, 2)

    def test_named_arguments_are_sent_as_named_over_the_wire(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        sheets = proxy.Worksheets()

        sheets.Add(After=sheets)

        method, params = client.calls[-1]
        self.assertEqual(method, 'invoke')
        self.assertEqual(params['name'], 'Add')
        self.assertEqual(params['args'], [])
        self.assertEqual(params['named'], {'After': {'$ole_ref': 2}})

    def test_property_set_uses_setattr(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        proxy.Visible = False

        method, params = client.calls[-1]
        self.assertEqual(method, 'invoke')
        self.assertEqual(params['name'], 'Visible=')
        self.assertEqual(params['args'], [False])

    def test_default_indexer_uses_empty_name(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        sheets = proxy.Worksheets()

        sheets[1]

        method, params = client.calls[-1]
        self.assertEqual(params['name'], '')
        self.assertEqual(params['args'], [1])

    def test_ole_const_load_returns_the_raw_dict(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        consts = proxy.ole_const_load()

        self.assertEqual(consts, {'xlUp': -4162, 'xlDown': -4121})

    def test_proxy_create_reports_created_true(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        self.assertTrue(proxy.ole_created)

    def test_proxy_connect_reports_created_false(self):
        client = FakeClient()
        proxy = Proxy.connect('Excel.Application', client)
        self.assertFalse(proxy.ole_created)

    def test_proxy_connect_or_create_reports_created_from_the_wire(self):
        client = FakeClient()
        client.connect_or_create_created = False
        proxy = Proxy.connect_or_create('Excel.Application', client)

        self.assertFalse(proxy.ole_created)
        method, params = client.calls[-1]
        self.assertEqual(method, 'connect_or_create')
        self.assertEqual(params['class_name'], 'Excel.Application')

    def test_proxy_wrap_reports_created_as_none(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        sheets = proxy.Worksheets()

        self.assertIsNone(sheets.ole_created)

    def test_invoke_is_a_public_escape_hatch(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        result = proxy.invoke('Version', [], {})

        self.assertEqual(result, 11.0)
        method, params = client.calls[-1]
        self.assertEqual(method, 'invoke')
        self.assertEqual(params['name'], 'Version')

    def test_decodes_tagged_time_into_a_python_datetime(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        result = proxy.Timestamp()

        self.assertIsInstance(result, datetime.datetime)
        self.assertEqual(result.year, 2026)

    def test_pickling_raises_not_serializable(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        with self.assertRaises(NotSerializableError):
            pickle.dumps(proxy)

    def test_stale_reference_raises_without_calling_the_client(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        other_client = FakeClient()
        other_client.generation = 999
        stale = Proxy.wrap(client, other_client.generation, 1)

        with self.assertRaises(StaleReferenceError):
            stale.Version()
        self.assertEqual(len(client.calls), 1)  # only the initial 'create' call

    def test_proxy_from_a_different_client_cannot_be_passed_as_an_argument(self):
        client_a = FakeClient()
        client_b = FakeClient()
        proxy_a = Proxy.create('Excel.Application', client_a)
        proxy_b = Proxy.create('Excel.Application', client_b)
        sheets_a = proxy_a.Worksheets()

        with self.assertRaises(StaleReferenceError):
            sheets_a.Add(After=proxy_b)

    def test_implicit_conversions_do_not_reach_the_rpc_layer(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        calls_before = len(client.calls)

        # These use CPython's type-level special-method lookup, which never
        # consults __getattr__ for a class that doesn't define the dunder --
        # so none of them should touch the RPC layer at all.
        self.assertTrue(bool(proxy))
        with self.assertRaises(TypeError):
            len(proxy)
        with self.assertRaises(TypeError):
            int(proxy)

        # copy.deepcopy is different: it probes the *instance* with a plain
        # getattr(obj, '__deepcopy__', None), which does reach __getattr__.
        # It must not become an RPC for a bogus '__deepcopy__' COM member --
        # it must fall through to __reduce__ and raise NotSerializableError.
        with self.assertRaises(NotSerializableError):
            copy.deepcopy(proxy)

        # Proxy defines no __iter__ of its own beyond an explicit refusal,
        # so without one iter()/for/`in` would fall back to the legacy
        # 0-based __getitem__ sequence protocol and fire one RPC per index.
        with self.assertRaises(TypeError):
            iter(proxy)
        with self.assertRaises(TypeError):
            list(proxy)

        self.assertEqual(len(client.calls), calls_before)

    def test_decode_converts_values_nested_inside_a_list(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        rows = proxy.BulkValue()

        self.assertEqual(rows[0][0], 1)
        self.assertIsInstance(
            rows[0][1], datetime.datetime,
            'a date inside a bulk range read must decode to a datetime, not stay a raw dict',
        )
        self.assertEqual(rows[0][1], datetime.datetime(2026, 8, 31, 9, 30, 0))
        self.assertIsNone(rows[1][0])
        self.assertEqual(rows[1][1], 'text')

    def test_decode_converts_ole_refs_nested_inside_a_list(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        sheets = proxy.BulkRefs()

        self.assertIsInstance(sheets[0], Proxy)
        self.assertEqual(sheets[0].ole_handle, 42)
        self.assertEqual(sheets[1].ole_handle, 43)

    def test_decode_converts_values_nested_inside_a_plain_dict(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        result = proxy.NestedHash()

        self.assertIsInstance(result['when'], datetime.datetime)

    def test_encode_converts_a_datetime_to_the_wire_tag(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        proxy.invoke('Value=', [datetime.datetime(2026, 8, 31, 9, 30, 0)], {})

        _method, params = client.calls[-1]
        self.assertEqual(
            [{'$type': 'time', 'iso8601': '2026-08-31T09:30:00'}],
            params['args'],
            'a datetime must go out as the same tag the receive side emits',
        )

    def test_encode_converts_datetimes_nested_inside_a_list(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        proxy.invoke('Value=', [[[datetime.datetime(2026, 8, 31, 9, 30, 0), 1]]], {})

        _method, params = client.calls[-1]
        self.assertEqual(
            [[[{'$type': 'time', 'iso8601': '2026-08-31T09:30:00'}, 1]]],
            params['args'],
            'encode is recursive, so a date inside a bulk write must be tagged too',
        )

    def test_encode_converts_a_date_to_the_wire_tag_at_midnight(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        proxy.invoke('Value=', [datetime.date(2026, 8, 31)], {})

        _method, params = client.calls[-1]
        self.assertEqual(
            [{'$type': 'time', 'iso8601': '2026-08-31T00:00:00'}],
            params['args'],
            'a bare date must go out as midnight',
        )

    # --- cleanup: (data-only) and ole_leave_open ---------------------------

    def test_create_sends_steps_and_callback_false(self):
        captured = {}

        class FakeCleanupClient:
            generation = 1

            def call(self, method, params):
                captured['method'] = method
                captured['params'] = params
                return {'$ole_ref': 3}

        Proxy.create('Excel.Application', FakeCleanupClient(),
                     cleanup={'steps': [['DisplayAlerts=', False], ['Quit']]})

        self.assertEqual(captured['method'], 'create')
        self.assertEqual(captured['params']['class_name'], 'Excel.Application')
        self.assertEqual(captured['params']['cleanup'], {
            'steps': [{'name': 'DisplayAlerts=', 'args': [False]}, {'name': 'Quit', 'args': []}],
            # Python has no client-side $cleanup closure path (no COM-event
            # delivery yet) -- callback must always be False, never True.
            'callback': False,
        })

    def test_connect_sends_steps_and_callback_false(self):
        captured = {}

        class FakeCleanupClient:
            generation = 1

            def call(self, method, params):
                captured['method'] = method
                captured['params'] = params
                return {'$ole_ref': 3}

        Proxy.connect('Excel.Application', FakeCleanupClient(), cleanup={'steps': [['Quit']]})

        self.assertEqual(captured['method'], 'connect')
        self.assertEqual(captured['params']['cleanup'], {
            'steps': [{'name': 'Quit', 'args': []}],
            'callback': False,
        })

    def test_connect_or_create_sends_steps_and_callback_false(self):
        client = FakeClient()

        Proxy.connect_or_create('Excel.Application', client, cleanup={'steps': [['Quit']]})

        _method, params = client.calls[-1]
        self.assertEqual(params['cleanup'], {
            'steps': [{'name': 'Quit', 'args': []}],
            'callback': False,
        })

    def test_create_without_a_cleanup_argument_sends_no_cleanup_key(self):
        client = FakeClient()

        Proxy.create('Excel.Application', client)

        _method, params = client.calls[-1]
        self.assertNotIn('cleanup', params, 'a caller that never asked for cleanup must not get a cleanup key on the wire')

    def test_ole_leave_open_sends_the_wire_call_and_returns_none(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        result = proxy.ole_leave_open()

        self.assertIsNone(result)
        method, params = client.calls[-1]
        self.assertEqual(method, 'leave_open')
        self.assertEqual(params['handle'], proxy.ole_handle)

    # --- on_cleanup, ole_release's cleanup-key reply, and ole_events -------

    def test_create_sends_callback_true_and_registers_the_closure(self):
        client = FakeClient()

        def on_cleanup():
            pass

        # on_cleanup present => callback True; steps [name, *args] =>
        # {name, args}. The closure itself never goes on the wire.
        proxy = Proxy.create('Excel.Application', client, cleanup={
            'steps': [['DisplayAlerts=', False], ['Quit']],
            'on_cleanup': on_cleanup,
        })

        method, params = client.calls[0]
        self.assertEqual(method, 'create')
        self.assertEqual(params['class_name'], 'Excel.Application')
        self.assertEqual(params['cleanup'], {
            'steps': [{'name': 'DisplayAlerts=', 'args': [False]}, {'name': 'Quit', 'args': []}],
            'callback': True,
        })
        self.assertEqual(client.dispatcher.registered_cleanups, [(proxy.ole_handle, on_cleanup)],
                         'the client closure must be registered against the created handle')

    def test_create_without_on_cleanup_sends_callback_false_and_registers_nothing(self):
        client = FakeClient()

        Proxy.create('Excel.Application', client, cleanup={'steps': [['Quit']]})

        _method, params = client.calls[0]
        self.assertEqual(params['cleanup'],
                         {'steps': [{'name': 'Quit', 'args': []}], 'callback': False})
        self.assertEqual(client.dispatcher.registered_cleanups, [],
                         'nothing to run means nothing to register, and callback stays False')

    def test_ole_release_returns_none(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        self.assertIsNone(proxy.ole_release())

    def test_ole_release_awaits_cleanup_when_the_bridge_replies_with_a_cleanup_key(self):
        client = FakeClient()
        client.release_reply = {'cleanup': 5}
        proxy = Proxy.create('Excel.Application', client)

        self.assertIsNone(proxy.ole_release())
        self.assertEqual(client.await_cleanup_calls, [5],
                         'a client closure must run before the handle is actually gone')

    def test_ole_release_does_not_await_when_the_bridge_replies_with_no_cleanup_key(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        proxy.ole_release()

        self.assertEqual(client.await_cleanup_calls, [],
                         'without a cleanup seq there is nothing to wait for')

    def test_ole_events_is_memoised_and_refused_on_a_stale_session(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)
        calls_before = len(client.calls)

        events = proxy.ole_events

        self.assertIsInstance(events, Events)
        self.assertIs(events, proxy.ole_events,
                      'memoised: a fresh Events per call would mean `on` and the `off` meant to '
                      'undo it talked to different objects')
        self.assertIs(events.proxy, proxy, 'and it must know the Proxy it belongs to')
        self.assertEqual(len(client.calls), calls_before,
                         'merely touching ole_events must cost no round trip')

        stale = Proxy.wrap(client, client.generation + 999, 1)
        with self.assertRaises(StaleReferenceError):
            stale.ole_events


if __name__ == '__main__':
    unittest.main()
