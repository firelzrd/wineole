import os
import signal
import subprocess
import sys
import tempfile
import time
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import wineole
from wineole.client import Client


def excel_pids():
    try:
        out = subprocess.run(
            ['pgrep', '-f', 'EXCEL.EXE'],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, check=False,
        ).stdout
        return {int(pid) for pid in out.split()}
    except OSError:
        return set()


def excel_gone(pre_existing_pids, timeout):
    deadline = time.monotonic() + timeout
    while True:
        if not (excel_pids() - pre_existing_pids):
            return True
        if time.monotonic() > deadline:
            return False
        time.sleep(0.5)


class WineOLEIntegrationTest(unittest.TestCase):
    def test_end_to_end_excel_automation_via_the_bridge(self):
        bridge_exe = Client.default_bridge_path()
        if not os.path.exists(bridge_exe):
            self.skipTest(f"bridge exe not built: {bridge_exe}")

        # Below 32768: Linux's default ephemeral port range is 32768-60999,
        # so a pid-derived port up there can collide with an outbound
        # connection the kernel handed out to some unrelated process.
        port = 20000 + (os.getpid() % 1000)
        spawned = {}

        # Excel is started by COM activation, not by us, so it has no PID we
        # can capture at spawn time. Snapshot what was already running
        # instead, so cleanup can kill only what this run caused — never
        # someone else's Excel, possibly with unsaved work, as a blanket
        # `pkill -f EXCEL.EXE` would.
        pre_existing_excel_pids = excel_pids()

        def spawner(p):
            proc = subprocess.Popen(
                ['wine', bridge_exe, str(p)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            spawned['proc'] = proc

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                try:
                    client = wineole.open(
                        port=port, spawner=spawner, lockfile=os.path.join(tmpdir, 'lock'), timeout=20
                    )

                    try:
                        xl = wineole.create('Excel.Application')
                        xl.Visible = False
                        xl.DisplayAlerts = False
                        xl.Workbooks().Add()

                        sheets = xl.Worksheets()
                        before_count = sheets.Count()
                        first_sheet = sheets[1]
                        first_sheet_name = first_sheet.Name()

                        consts = xl.ole_const_load()
                        self.assertEqual(consts.get('xlUp'), -4162)

                        new_sheet = sheets.Add(After=first_sheet)
                        after_count = xl.Worksheets().Count()

                        self.assertEqual(after_count, before_count + 1)
                        # Position, not just count — mirrors the Ruby
                        # integration test's own fix: prove the named
                        # argument was honored, not merely accepted.
                        self.assertEqual(xl.Worksheets()[1].Name(), first_sheet_name)
                        self.assertNotEqual(xl.Worksheets()[2].Name(), first_sheet_name)
                        # Prove it's genuinely the new sheet at index 2, not
                        # just "something different from first_sheet" (which
                        # a misapplied append-at-the-end would also satisfy
                        # in a workbook with more than 2 sheets).
                        self.assertEqual(xl.Worksheets()[2].Name(), new_sheet.Name())

                        xl.Quit()
                    finally:
                        client.close()
                        wineole._default_client = None
                finally:
                    # This must run even if open() itself raised
                    # (e.g. WineOLEError on timeout, or a re-raised handshake
                    # failure) — otherwise a process spawner already put into
                    # `spawned` is orphaned, leaking a wine process.
                    if 'proc' in spawned:
                        proc = spawned['proc']
                        try:
                            # SIGTERM, not SIGKILL: gives the bridge a chance
                            # to run its own session/COM teardown, which is
                            # what releases Excel cleanly. Mirrors the Ruby
                            # test's explicit choice of TERM for the same
                            # reason.
                            proc.terminate()
                        except ProcessLookupError:
                            pass
                        try:
                            proc.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            # A bridge that ignores or is wedged past SIGTERM
                            # must still be reaped — otherwise the wine
                            # process outlives the test run.
                            proc.kill()
                            try:
                                proc.wait(timeout=5)
                            except subprocess.TimeoutExpired:
                                pass
        finally:
            # Kill only what this run started, by PID. A blanket
            # `pkill -f EXCEL.EXE` would kill every Excel on the machine,
            # including unrelated ones with unsaved work.
            excel_gone(pre_existing_excel_pids, timeout=5)
            for pid in excel_pids() - pre_existing_excel_pids:
                try:
                    os.kill(pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass

    def test_connect_or_create_creates_then_a_second_call_attaches(self):
        bridge_exe = Client.default_bridge_path()
        if not os.path.exists(bridge_exe):
            self.skipTest(f"bridge exe not built: {bridge_exe}")

        port = 20000 + (os.getpid() % 1000)
        spawned = {}
        pre_existing_excel_pids = excel_pids()

        def spawner(p):
            proc = subprocess.Popen(
                ['wine', bridge_exe, str(p)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
            spawned['proc'] = proc

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                try:
                    client = wineole.open(
                        port=port, spawner=spawner, lockfile=os.path.join(tmpdir, 'lock'), timeout=20
                    )
                    try:
                        xl = client.connect_or_create('Excel.Application')
                        self.assertTrue(xl.ole_created, 'the first connect_or_create must create a new instance')
                        xl.Visible = False
                        xl.DisplayAlerts = False

                        xl2 = client.connect_or_create('Excel.Application')
                        self.assertFalse(xl2.ole_created, 'the second connect_or_create must attach to the instance the first one created')

                        # Prove it's genuinely the same live instance, not a
                        # coincidentally similar second one.
                        xl2.DisplayAlerts = True
                        xl.DisplayAlerts = False
                        self.assertFalse(xl2.DisplayAlerts(), 'xl and xl2 must observe the same live Excel instance')

                        xl.Quit()
                    finally:
                        client.close()
                finally:
                    if 'proc' in spawned:
                        proc = spawned['proc']
                        try:
                            proc.terminate()
                        except ProcessLookupError:
                            pass
                        try:
                            proc.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            proc.kill()
                            try:
                                proc.wait(timeout=5)
                            except subprocess.TimeoutExpired:
                                pass
        finally:
            excel_gone(pre_existing_excel_pids, timeout=5)
            for pid in excel_pids() - pre_existing_excel_pids:
                try:
                    os.kill(pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass


if __name__ == '__main__':
    unittest.main()
