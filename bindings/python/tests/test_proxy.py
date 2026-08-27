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

_fake_client_generation = itertools.count(1)


class FakeClient:
    def __init__(self):
        self.calls = []
        # A fresh, distinct generation per instance — mirroring
        # wineole/client.py's Client.generation (an explicit per-instance
        # counter, deliberately not id()). A hardcoded `self.generation = 1`
        # here would make every FakeClient indistinguishable from every
        # other, silently defeating any test that checks cross-client
        # behavior (two different clients must never compare equal).
        self.generation = next(_fake_client_generation)
        self.connect_or_create_created = True

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
            return None
        if method == 'invoke':
            name = params['name']
            if name == 'Version':
                return 11.0
            if name == 'Worksheets':
                return {'$ole_ref': 2}
            if name == 'Timestamp':
                return {'$type': 'time', 'iso8601': '2026-08-26T04:00:00+00:00'}
            return None
        return None


class ProxyTest(unittest.TestCase):
    def test_attribute_access_returns_a_callable_not_an_immediate_value(self):
        client = FakeClient()
        proxy = Proxy.create('Excel.Application', client)

        member = proxy.Version  # bare attribute access — must NOT call the client yet
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
        # consults __getattr__ for a class that doesn't define the dunder —
        # so none of them should touch the RPC layer at all.
        self.assertTrue(bool(proxy))
        with self.assertRaises(TypeError):
            len(proxy)
        with self.assertRaises(TypeError):
            int(proxy)

        # copy.deepcopy is different: it probes the *instance* with a plain
        # getattr(obj, '__deepcopy__', None), which does reach __getattr__.
        # It must not become an RPC for a bogus '__deepcopy__' COM member —
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


if __name__ == '__main__':
    unittest.main()
