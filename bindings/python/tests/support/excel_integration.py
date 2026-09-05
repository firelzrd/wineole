"""Shared spawn/teardown plumbing for the real-Excel integration tests.

A port of the private helpers inside tests/test_integration.py, lifted here
so the msoffice integration file does not re-copy the spawn/cleanup dance --
getting one copy of it subtly wrong is how a stray Excel process survives a
test run. test_integration.py itself is deliberately left alone for now; see
the plan's Task 8 note.
"""

import contextlib
import os
import signal
import subprocess
import tempfile
import threading
import time

import wineole
from wineole.client import Client


def excel_pids():
    """The Excel processes running right now.

    `-x` matches the process name exactly; `-f` would match any command line
    that merely mentions the executable (a diagnostic pgrep, a filesystem
    search for the file), which silently skipped a test that requires "no
    Excel running" and, worse, handed teardown an unrelated process to kill.
    """
    try:
        out = subprocess.run(
            ['pgrep', '-x', 'EXCEL.EXE'],
            stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, text=True, check=False,
        ).stdout
        return {int(pid) for pid in out.split()}
    except OSError:
        return set()


def excel_gone(pre_existing_pids, timeout):
    """Checking for a leftover Excel immediately after a run is racy: the
    process can be mid-teardown (Wine's own COM release/exit sequence) and
    gone moments later, so poll with a short timeout rather than sampling
    once."""
    deadline = time.monotonic() + timeout
    while True:
        if not (excel_pids() - pre_existing_pids):
            return True
        if time.monotonic() > deadline:
            return False
        time.sleep(0.5)


class ExcelIntegrationMixin:
    """Mixed into a unittest.TestCase. `self.pre_existing_excel_pids` is the
    snapshot taken when a bridge comes up; teardown kills only what is not
    in it."""

    @contextlib.contextmanager
    def bridge(self):
        """Bring up a bridge and yield a client; tear the bridge down after.

        Split from `raw_excel` so a test that needs only a client -- to hand
        to Excel.create or Excel.run -- can reuse the spawn/teardown logic
        without a create() call defeating its "nothing exists yet"
        precondition.
        """
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
        # instead, so cleanup can kill only what this run caused -- never
        # someone else's Excel, possibly with unsaved work.
        self.pre_existing_excel_pids = excel_pids()

        def spawner(p):
            spawned['proc'] = subprocess.Popen(
                ['wine', bridge_exe, str(p)],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )

        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                try:
                    client = wineole.open(
                        port=port, spawner=spawner,
                        lockfile=os.path.join(tmpdir, 'lock'), timeout=20,
                    )
                    try:
                        yield client
                    finally:
                        wineole.close()
                finally:
                    # This must run even if open() itself raised, otherwise
                    # a spawned process is orphaned and a wine process leaks.
                    proc = spawned.get('proc')
                    if proc is not None:
                        try:
                            # SIGTERM, not SIGKILL: gives the bridge a chance
                            # to run its own session/COM teardown, which is
                            # what releases Excel cleanly.
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
            # Kill only what this run started, by PID.
            excel_gone(self.pre_existing_excel_pids, timeout=5)
            for pid in excel_pids() - self.pre_existing_excel_pids:
                try:
                    os.kill(pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass

    @contextlib.contextmanager
    def raw_excel(self):
        """A bridge plus a raw COM Application, both torn down after.

        Quit is attempted best-effort on the way out, so `bridge()`'s
        kill-by-PID cleanup stays a safety net rather than the primary way
        Excel goes away.
        """
        with self.bridge():
            xl = wineole.create('Excel.Application')
            try:
                yield xl
            finally:
                self.quit_bounded(xl)

    def quit_bounded(self, xl, seconds=20):
        """Quit on a connection the test body may have wedged, with a clock
        on it. An unbounded Quit turns a failing test into a hanging suite --
        no output, no result, and every later test taken with it."""
        done = threading.Event()

        def quit_it():
            try:
                xl.Quit()
            except Exception:
                pass
            finally:
                done.set()

        threading.Thread(target=quit_it, daemon=True).start()
        done.wait(seconds)
