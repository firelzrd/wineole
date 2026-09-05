import contextlib
import datetime
import os
import queue
import signal
import subprocess
import sys
import tempfile
import threading
import time
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import wineole
from wineole.client import Client


def excel_pids():
    try:
        # -x matches the process name exactly; -f would match any command
        # line that merely mentions EXCEL.EXE (a diagnostic pgrep, a
        # filesystem search for the file), which silently skipped a test
        # that requires "no Excel running" and, worse, handed teardown an
        # unrelated process to kill.
        out = subprocess.run(
            ['pgrep', '-x', 'EXCEL.EXE'],
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
    @contextlib.contextmanager
    def _bridge(self):
        """Bring up a bridge and yield a client; tear the bridge down after.

        Split out from _excel() so
        test_connect_or_create_creates_then_a_second_call_attaches can reuse
        the spawn/cleanup dance without _excel()'s wineole.create() call
        defeating its "nothing exists yet" precondition. _excel() below adds
        only the Excel-specific bring-up/teardown on top of this.
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
                        yield client
                    finally:
                        wineole.close()
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

    @contextlib.contextmanager
    def _excel(self):
        """Bring up a bridge and an Excel instance, and tear both down.

        Lifted out of test_end_to_end_excel_automation_via_the_bridge so the
        bulk-array, latency and shape-contract tests do not each re-copy the
        spawn/cleanup dance -- getting one copy of it subtly wrong is how a
        stray EXCEL.EXE survives a test run.

        Mirrors the Ruby test's with_excel: Quit is attempted best-effort on
        the way out, so _bridge()'s kill-by-PID cleanup stays a safety net
        rather than the primary way Excel goes away.
        """
        with self._bridge():
            xl = wineole.create('Excel.Application')
            try:
                yield xl
            finally:
                try:
                    xl.Quit()
                except Exception:
                    pass

    def _quit_bounded(self, xl, seconds=20):
        """Quit on a connection the test body may have wedged, with a clock on
        it. An unbounded Quit turns a failing test into a hanging suite -- no
        output, no result, and every later test taken with it. Measured on the
        Ruby events test that exists to catch callbacks running on the reader
        thread: under that mutation the run produced nothing for 200s and had
        to be killed, against 1.6s for a correct answer."""
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

    def test_end_to_end_excel_automation_via_the_bridge(self):
        with self._excel() as xl:
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

    def test_connect_or_create_creates_then_a_second_call_attaches(self):
        # Deliberately using _bridge() rather than _excel(): _excel() brings
        # up its xl via wineole.create(), which would itself already be the
        # "first" instance and defeat this test's point -- that the *first*
        # connect_or_create() call is what creates one. This test needs a
        # from-scratch bring-up so nothing exists yet when it calls
        # connect_or_create() the first time.
        #
        # Not `excel_pids() - pre_existing_excel_pids` -- there is no fresh
        # snapshot taken right before this test, so a leftover Excel from a
        # sibling test would be captured *into* any such snapshot and the
        # difference would always be empty. connect_or_create attaches to
        # ANY running Excel.Application, test-started or not, so the only
        # clean precondition is that none is running at all.
        if excel_pids():
            self.skipTest(
                'an Excel instance is already running -- cannot assert "nothing was running" cleanly; '
                're-run this test in isolation or investigate the leftover instance'
            )

        with self._bridge() as client:
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

    def test_bulk_range_round_trip_preserves_types(self):
        with self._excel() as xl:
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.Workbooks().Add()
            sheet = xl.Worksheets()[1]

            # One write for the whole block, not nine.
            sheet.Range('A1:C3').Value = [
                ['text', 1, 2.5],
                ['', None, -3],
                ['ünïcödé ✓', 0, 1000000],
            ]

            rows = sheet.Range('A1:C3').Value()

            self.assertEqual(len(rows), 3, 'a 3x3 range must read back as 3 rows')
            self.assertEqual(len(rows[0]), 3, 'each row must have 3 columns (not transposed)')
            self.assertEqual(rows[0][0], 'text')
            self.assertEqual(rows[0][1], 1.0)
            self.assertEqual(rows[0][2], 2.5)
            # A None written into a cell comes back as None.
            self.assertIsNone(rows[1][1], 'a None written into a cell must read back as None')
            # An empty string written into a cell comes back as an empty
            # cell (None), not as ''.
            self.assertIsNone(rows[1][0], "an empty string written into a cell must read back as None, not ''")
            self.assertEqual(rows[1][2], -3.0)
            self.assertEqual(rows[2][0], 'ünïcödé ✓')
            self.assertEqual(rows[2][2], 1000000.0)

            # A date cell must arrive as a datetime, not as a raw
            # {'$type': 'time'} dict -- this is what the recursive decode
            # exists for.
            #
            # A 1x1 range's Value is a bare scalar, not [[v]] (see
            # test_range_value_shape_depends_on_range_size) -- so `date`
            # here is the datetime itself, not a row containing it. That
            # only proves *top-level* decode of a tagged value.
            sheet.Range('E1').Value = '2026-08-31'
            sheet.Range('E1').NumberFormat = 'yyyy-mm-dd'
            date = sheet.Range('E1:E1').Value()
            self.assertIsInstance(
                date, datetime.datetime,
                'a date inside a bulk read must decode to a datetime',
            )
            self.assertEqual(date.year, 2026)
            self.assertEqual(date.month, 8)
            self.assertEqual(date.day, 31)

            # The assertion above says nothing about a tagged value *nested*
            # inside a returned array -- which is exactly the case the
            # recursive decode exists for (a bulk Range.Value read on
            # anything larger than 1x1 returns a list of rows, per
            # test_range_value_shape_depends_on_range_size). Write two dates
            # side by side and read them back as a 1x2 range, so reaching
            # either date means indexing into the returned list:
            # [[datetime, datetime]], not a bare datetime.
            sheet.Range('F1').Value = '2026-09-01'
            sheet.Range('F1').NumberFormat = 'yyyy-mm-dd'
            dates = sheet.Range('E1:F1').Value()
            self.assertEqual(len(dates), 1, 'a 1x2 range must read back as 1 row')
            self.assertEqual(len(dates[0]), 2, 'that row must have 2 columns')
            self.assertIsInstance(
                dates[0][0], datetime.datetime,
                "a date nested inside a returned list must decode to a datetime, not a raw {'$type': 'time'} dict",
            )
            self.assertIsInstance(
                dates[0][1], datetime.datetime,
                "a date nested inside a returned list must decode to a datetime, not a raw {'$type': 'time'} dict",
            )
            self.assertEqual(dates[0][0].year, 2026)
            self.assertEqual(dates[0][0].month, 8)
            self.assertEqual(dates[0][0].day, 31)
            self.assertEqual(dates[0][1].year, 2026)
            self.assertEqual(dates[0][1].month, 9)
            self.assertEqual(dates[0][1].day, 1)

    def test_writing_a_datetime_directly_round_trips_as_a_date(self):
        with self._excel() as xl:
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.Workbooks().Add()
            sheet = xl.Worksheets()[1]

            # A datetime assigned straight into a cell -- no
            # string-plus-NumberFormat workaround -- must land as a genuine
            # VT_DATE and read back as a datetime. This asserts on the
            # *type* read back rather than on display text: Excel's own
            # Range.Value setter auto-applies a date format to a
            # still-General cell (see README), so a passing string-based
            # workaround and a genuine VT_DATE can look identical on
            # screen -- only the read-back type tells them apart.
            written = datetime.datetime(2026, 8, 31, 9, 30, 45)
            sheet.Range('H1').Value = written
            read_back = sheet.Range('H1:H1').Value()
            self.assertIsInstance(
                read_back, datetime.datetime,
                'a datetime written straight into a cell must read back as a datetime, not a str or a raw dict',
            )
            self.assertEqual(
                int(written.timestamp()), int(read_back.timestamp()),
                'a direct datetime write must round-trip to the second',
            )

            # The same, but nested inside a 2-D bulk write -- the recursive
            # case in value.rs's SAFEARRAY encode/decode, which the
            # single-cell write above does not exercise. Mixed with plain
            # strings so the test also proves the recursion doesn't
            # misencode/misdecode a date's neighbors.
            written_a = datetime.datetime(2026, 9, 1, 12, 0, 0)
            written_b = datetime.datetime(2026, 9, 2, 18, 15, 30)
            sheet.Range('I1:J2').Value = [
                [written_a, 'not a date'],
                ['still not a date', written_b],
            ]
            grid = sheet.Range('I1:J2').Value()

            self.assertIsInstance(
                grid[0][0], datetime.datetime,
                'a date written inside a bulk 2-D array must read back as a datetime, indexed out of the returned list',
            )
            self.assertIsInstance(
                grid[1][1], datetime.datetime,
                'a date written inside a bulk 2-D array must read back as a datetime, indexed out of the returned list',
            )
            self.assertEqual(
                int(written_a.timestamp()), int(grid[0][0].timestamp()),
                'a bulk-written date must round-trip to the second',
            )
            self.assertEqual(
                int(written_b.timestamp()), int(grid[1][1].timestamp()),
                'a bulk-written date must round-trip to the second',
            )
            self.assertEqual(grid[0][1], 'not a date')
            self.assertEqual(grid[1][0], 'still not a date')

    def test_round_trip_latency_is_not_stalled_by_nagle(self):
        with self._excel() as xl:
            xl.Version()  # warm up: the first call also starts Excel

            started = time.monotonic()
            for _ in range(100):
                xl.Version()
            per_call_ms = (time.monotonic() - started) / 100 * 1000

            # Before the single-write fix this was 42.6 ms per call,
            # essentially all of it the client's ~40 ms delayed-ACK timer
            # waiting for a newline Nagle was holding. Afterwards it is
            # ~1.3 ms. 20 ms sits clear of both, so this catches a regression
            # without being flaky on a loaded machine.
            self.assertLess(
                per_call_ms, 20.0,
                f'a round trip took {per_call_ms:.1f} ms; a Nagle/delayed-ACK stall '
                'has probably come back (see protocol.rs write_response)',
            )

    # --- cleanup: steps and leave_open (Task 10, data-only) ----------------
    #
    # Python has no COM-event delivery yet, so it never sets callback: true
    # -- these tests only exercise the STEPS path: the bridge running the
    # declared steps itself when the instance's last user releases the root,
    # and leave_open revoking that permission. Deliberately using _bridge()
    # rather than _excel(): _excel() creates its own Excel via
    # wineole.create() with no cleanup, which would be a second,
    # independent instance muddying "did *this* create's cleanup quit
    # Excel" -- and _excel()'s teardown calls xl.Quit() on its own instance,
    # not the one under test here.
    def test_cleanup_steps_quit_excel_when_the_last_user_releases_it(self):
        with self._bridge() as client:
            pre_existing_excel_pids = excel_pids()

            xl = client.create('Excel.Application',
                               cleanup={'steps': [['DisplayAlerts=', False], ['Quit']]})
            xl.Visible = False
            xl.Workbooks().Add()

            xl.ole_release()

            self.assertTrue(
                excel_gone(pre_existing_excel_pids, timeout=10),
                'the declared cleanup steps must quit the auto-created Excel once its last user releases it',
            )

    def test_leave_open_keeps_excel_running_after_release(self):
        with self._bridge() as client:
            pre_existing_excel_pids = excel_pids()

            xl = client.create('Excel.Application',
                               cleanup={'steps': [['DisplayAlerts=', False], ['Quit']]})
            xl.Visible = False
            xl.Workbooks().Add()

            xl.ole_leave_open()
            xl.ole_release()

            # excel_gone would return True immediately if nothing new ever
            # ran; give the (revoked) cleanup a real chance to fire wrongly
            # before asserting the instance survived it. _bridge()'s own
            # teardown (above, in its finally block) kills whatever is left
            # by PID once this `with` exits -- no need to repeat that here.
            self.assertFalse(
                excel_gone(pre_existing_excel_pids, timeout=5),
                'ole_leave_open must revoke the cleanup steps -- Excel must still be running after release',
            )

    def test_range_value_shape_depends_on_range_size(self):
        # Range.Value's result shape depends on the range's dimensions --
        # this is Excel's own contract, not a choice this project made, and
        # Phase 2's wrapper has to normalize it. Pin it here so a future
        # change to the conversion layer cannot quietly alter it out from
        # under that wrapper.
        with self._excel() as xl:
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.Workbooks().Add()
            sheet = xl.Worksheets()[1]

            sheet.Range('A1').Value = 42
            sheet.Range('B1').Value = 'right'
            sheet.Range('A2').Value = 'below'

            # A 1x1 range collapses to a bare scalar, whether addressed as a
            # single cell or as an explicit 1x1 range -- not [[42.0]].
            self.assertEqual(sheet.Range('A1').Value(), 42.0)
            self.assertEqual(sheet.Range('A1:A1').Value(), 42.0)

            # Anything larger than 1x1 comes back as a list of row lists.
            self.assertEqual(sheet.Range('A1:B1').Value(), [[42.0, 'right']])
            self.assertEqual(sheet.Range('A1:A2').Value(), [[42.0], ['below']])

            # An empty 1x1 range is a bare None, not [[None]].
            self.assertIsNone(
                sheet.Range('J1:J1').Value(),
                'an empty 1x1 range must read back as a bare None, not [[None]]',
            )
            self.assertEqual(sheet.Range('J1:J2').Value(), [[None], [None]])

    # --- COM events and on_cleanup, against a real Excel -------------------

    # THE regression test for the thread split. A callback that calls COM must
    # get an answer. Running callbacks on the reader thread makes this hang
    # forever, because nothing is left to read the response.
    def test_a_callback_can_call_com_and_get_an_answer(self):
        with self._bridge() as client:
            xl = client.create('Excel.Application')
            try:
                xl.Visible = False
                xl.DisplayAlerts = False
                got = queue.Queue()

                def on_change(sheet, target):
                    try:
                        got.put((sheet.Name(), target.Address()))
                    except Exception as exc:
                        got.put(exc)

                xl.ole_events.on('SheetChange', on_change)
                book = xl.Workbooks().Add()
                book.Worksheets()[1].Range('A1').Value = 42

                try:
                    seen = got.get(timeout=30)
                except queue.Empty:
                    # A verdict, not an exception escaping: this is the
                    # failure this test exists to report, and saying so is the
                    # difference between a suite that tells you what broke and
                    # one that times out.
                    self.fail('the callback never got an answer within 30s -- with callbacks '
                              'running on the reader thread there is nobody left to read the '
                              "response to the callback's own COM call, and the whole connection "
                              'is wedged')
                self.assertNotIsInstance(seen, Exception, f"the callback raised: {seen!r}")
                name, address = seen
                self.assertRegex(address, r'\$?A\$?1',
                                 'the callback made a COM call and got an answer back')
                self.assertTrue(name, 'and so did the second one')
            finally:
                self._quit_bounded(xl)

    # `on` is the only thing the caller touches. subscribe and Advise are
    # derived; removing that derivation makes this fail. And `off` is the
    # other half of the same claim: the write still raises SheetChange in
    # Excel, and nothing reaches the client because the last callback for it
    # is gone.
    def test_registering_a_callback_is_all_it_takes(self):
        with self._bridge() as client:
            xl = client.create('Excel.Application')
            try:
                xl.Visible = False
                xl.DisplayAlerts = False
                fired = queue.Queue()
                sub = xl.ole_events.on('SheetChange', lambda *args: fired.put('changed'))

                book = xl.Workbooks().Add()
                sheet = book.Worksheets()[1]
                sheet.Range('B2').Value = 1

                try:
                    self.assertEqual(fired.get(timeout=30), 'changed')
                except queue.Empty:
                    self.fail('no event arrived within 30s: registering a callback must be all it '
                              'takes')

                xl.ole_events.off(sub)
                sheet.Range('B3').Value = 2
                # Bounded, and it must fire before the off or this proves
                # nothing -- which the assertion above is what establishes.
                try:
                    late = fired.get(timeout=5)
                except queue.Empty:
                    late = None
                self.assertIsNone(late, 'off means off: nothing may arrive after the last callback '
                                        'for the event is gone')
            finally:
                self._quit_bounded(xl)

    # One dispatcher thread per CONNECTION, on the real thing. Two objects
    # with callbacks on one client -- the Application and the Workbook, both
    # raising SheetChange for the same write -- and the promise is that a
    # caller who shares state between them needs no lock of his own, because
    # the two can never be inside their callbacks at once. A thread per object
    # breaks exactly that and nothing else.
    def test_callbacks_on_two_objects_run_on_the_one_dispatcher_thread(self):
        with self._bridge() as client:
            xl = client.create('Excel.Application')
            try:
                xl.Visible = False
                xl.DisplayAlerts = False
                book = xl.Workbooks().Add()

                fired = queue.Queue()
                xl.ole_events.on(
                    'SheetChange', lambda *args: fired.put(('application', threading.get_ident())))
                book.ole_events.on(
                    'SheetChange', lambda *args: fired.put(('workbook', threading.get_ident())))

                book.Worksheets()[1].Range('C3').Value = 7

                seen = {}
                deadline = time.monotonic() + 30
                while ('application' not in seen or 'workbook' not in seen):
                    left = deadline - time.monotonic()
                    if left <= 0:
                        # A verdict, not a timeout: with only one of the two
                        # delivered there is nothing to compare, and saying
                        # which one is missing is the difference between a
                        # report and a hang.
                        self.fail(f"only {sorted(seen)} fired within 30s -- both objects must be "
                                  'delivered to')
                    try:
                        who, ident = fired.get(timeout=left)
                    except queue.Empty:
                        continue
                    seen[who] = ident

                self.assertEqual(
                    seen['application'], seen['workbook'],
                    "both callbacks must run on the connection's ONE dispatcher thread: a thread "
                    "per object is a data race in every callback that shares state with another "
                    "object's")
            finally:
                self._quit_bounded(xl)

    # The closure is a prelude, not a veto. ole_release blocks until the
    # dispatcher has both run the closure and received the release_event
    # completion, so by the time it returns the flag the closure sets must
    # already be true -- and the steps still quit Excel afterwards. Whether
    # the process has finished exiting is a separate question, so that half
    # polls with a bounded wait rather than sampling once.
    def test_on_cleanup_closure_runs_then_steps_quit_excel(self):
        with self._bridge() as client:
            pre_existing_excel_pids = excel_pids()
            ran = []

            xl = client.create('Excel.Application', cleanup={
                'steps': [['DisplayAlerts=', False], ['Quit']],
                'on_cleanup': lambda: ran.append(True),
            })
            xl.Visible = False
            xl.Workbooks().Add()

            xl.ole_release()  # blocks until the $cleanup closure completes

            self.assertEqual(ran, [True],
                             'the on_cleanup closure must have run before ole_release returned')
            self.assertTrue(
                excel_gone(pre_existing_excel_pids, timeout=20),
                'the steps must still quit Excel after the closure runs')

    # A closure that calls ole_leave_open revokes the bridge's shutdown
    # permission from inside the callback itself, so the release that
    # triggered the closure runs no steps at all.
    def test_on_cleanup_closure_calling_leave_open_keeps_excel_open(self):
        with self._bridge() as client:
            pre_existing_excel_pids = excel_pids()
            root = {}

            xl = client.create('Excel.Application', cleanup={
                'steps': [['DisplayAlerts=', False], ['Quit']],
                # The closure only reads root['xl'] once the bridge actually
                # asks for it, on the dispatcher thread, by which point the
                # assignment below has happened.
                'on_cleanup': lambda: root['xl'].ole_leave_open(),
            })
            root['xl'] = xl
            xl.Visible = False
            xl.Workbooks().Add()

            xl.ole_release()

            # A grace period, not a formality: measured on Ruby with the
            # closure mutated to a no-op, ole_release still returned in ~2ms
            # (the bridge issues Quit synchronously before it unblocks) but
            # the process lingered for close to 1s past that before actually
            # exiting. Checking the instant ole_release returns would
            # therefore pass this assertion even with a broken
            # ole_leave_open -- a false positive from a process that is merely
            # still tearing down. Sleeping past that window first is what
            # makes "still running" mean the steps never ran at all.
            time.sleep(3)

            self.assertFalse(
                excel_gone(pre_existing_excel_pids, timeout=0),
                'a closure calling ole_leave_open must keep Excel running past ole_release')

            # Quit it here rather than leaving it to _bridge()'s kill-by-PID
            # safety net: a second connect on the same client attaches to the
            # instance the closure kept alive, which is also the proof that it
            # is still a live, reachable COM object and not just a process
            # that has not finished exiting.
            survivor = client.connect('Excel.Application')
            survivor.DisplayAlerts = False
            self._quit_bounded(survivor)


if __name__ == '__main__':
    unittest.main()
