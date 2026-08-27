import os
import sys
import threading
from unittest.mock import patch

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import unittest
import wineole


class FakeClient:
    def __init__(self, label):
        self.label = label

    def create(self, class_name):
        return f'{self.label}:{class_name}'

    def close(self):
        pass


class WineOLEModuleTest(unittest.TestCase):
    def tearDown(self):
        wineole.close()

    def test_default_client_lazy_init_is_thread_safe_and_calls_open_only_once(self):
        fake_client = FakeClient('shared')
        call_count = []
        count_lock = threading.Lock()

        def fake_open():
            with count_lock:
                call_count.append(1)
            threading.Event().wait(0.05)  # widen the race window
            return fake_client

        wineole.close()
        with patch.object(wineole.Client, 'open', staticmethod(fake_open)):
            results = []
            results_lock = threading.Lock()

            def worker():
                result = wineole.create('Excel.Application')
                with results_lock:
                    results.append(result)

            threads = [threading.Thread(target=worker) for _ in range(20)]
            for t in threads:
                t.start()
            for t in threads:
                t.join()

            self.assertEqual(len(call_count), 1, 'Client.open must be called exactly once no matter how many threads race')
            self.assertEqual(set(results), {'shared:Excel.Application'})

    def test_open_updates_the_implicit_default(self):
        fake_client_a = FakeClient('a')
        with patch.object(wineole.Client, 'open', staticmethod(lambda **kwargs: fake_client_a)):
            opened = wineole.open()
            self.assertIs(opened, fake_client_a)

            result = wineole.create('Excel.Application')
            self.assertEqual(result, 'a:Excel.Application',
                              'wineole.create must use the client wineole.open just set as the default')

    def test_close_clears_the_default_so_the_next_create_opens_a_fresh_one(self):
        clients = [FakeClient('a'), FakeClient('b')]

        def fake_open():
            return clients.pop(0)

        with patch.object(wineole.Client, 'open', staticmethod(fake_open)):
            wineole.create('Excel.Application')
            first_default = wineole._get_default_client()
            self.assertEqual(first_default.label, 'a')

            wineole.close()
            wineole.create('Excel.Application')
            second_default = wineole._get_default_client()

            self.assertEqual(second_default.label, 'b')
            self.assertIsNot(first_default, second_default,
                              'close() must force the next .create to open a fresh client')

    def test_connect_or_create_uses_the_default_client(self):
        class FakeClientWithConnectOrCreate(FakeClient):
            def connect_or_create(self, class_name):
                return f'coc:{class_name}'

        fake_client = FakeClientWithConnectOrCreate('unused')
        with patch.object(wineole.Client, 'open', staticmethod(lambda **kwargs: fake_client)):
            result = wineole.connect_or_create('Excel.Application')
            self.assertEqual(result, 'coc:Excel.Application')


if __name__ == '__main__':
    unittest.main()
