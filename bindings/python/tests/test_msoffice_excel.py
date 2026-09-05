import os
import subprocess
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import (
    FakeClient, FakeComApplication, FakeComWorkbook, FakeComWorksheet, SelectSpy,
)

import wineole
from wineole.msoffice.book import Book
from wineole.msoffice.excel import Excel
from wineole.msoffice.range import Range
from wineole.msoffice.sheet import Sheet


class Boom(Exception):
    pass


class MSOfficeExcelTest(unittest.TestCase):
    def setUp(self):
        SelectSpy.reset()

    def tearDown(self):
        self.assertEqual(
            [], SelectSpy.calls,
            'Excel must never call Select -- resolving through worksheet objects needs no '
            "active state, so a read must not mutate the caller's selection as a side effect")

    def new_excel(self, app, convert_paths=True):
        return Excel(app, FakeClient(app), convert_paths=convert_paths)

    # --- lifecycle: run always releases what it used ----------------------
    #
    # Whether a release also quits Excel is the bridge's decision (it runs
    # CLEANUP_STEPS only for the last user of a record it auto-created, per
    # the cleanup= the factories declare) -- not something this fake or
    # these tests can see. What `run` itself owns, and what these pin, is
    # that its finally calls ole_release exactly once no matter which mode
    # built the record.

    def test_run_releases_what_it_used_regardless_of_mode(self):
        created_app = FakeComApplication(created=True)
        with Excel.run('create', client=FakeClient(created_app)) as xl:
            self.assertIsInstance(xl, Excel)
        self.assertEqual(1, created_app.ole_release_calls)

        attached_app = FakeComApplication(created=False)
        with Excel.run('connect', client=FakeClient(attached_app)):
            pass
        self.assertEqual(1, attached_app.ole_release_calls)

    def test_run_with_connect_or_create_also_releases_regardless_of_the_bridges_report(self):
        freshly_created = FakeComApplication(created=True)
        with Excel.run('connect_or_create', client=FakeClient(freshly_created)):
            pass
        self.assertEqual(1, freshly_created.ole_release_calls)

        already_running = FakeComApplication(created=False)
        with Excel.run('connect_or_create', client=FakeClient(already_running)):
            pass
        self.assertEqual(1, already_running.ole_release_calls)

    def test_run_releases_even_when_the_block_raises(self):
        created_app = FakeComApplication(created=True)
        with self.assertRaises(Boom) as ctx:
            with Excel.run('create', client=FakeClient(created_app)):
                raise Boom('block failed')
        self.assertEqual('block failed', str(ctx.exception))
        self.assertEqual(1, created_app.ole_release_calls)

    def test_run_with_an_unknown_mode_raises_before_connecting(self):
        app = FakeComApplication(created=True)
        client = FakeClient(app)
        with self.assertRaises(ValueError) as ctx:
            Excel.run('attach', client=client)
        self.assertIn('unknown mode', str(ctx.exception))
        self.assertIn("'attach'", str(ctx.exception))
        self.assertEqual([], client.create_calls + client.connect_calls
                         + client.connect_or_create_calls,
                         'a bad mode must be refused before anything connects')

    # --- ole_release / leave_open delegate to the underlying proxy ---------

    def test_ole_release_delegates_to_the_proxy(self):
        a = FakeComApplication(created=True)
        self.new_excel(a).ole_release()
        self.assertEqual(1, a.ole_release_calls)

    def test_leave_open_delegates_to_the_proxy(self):
        a = FakeComApplication(created=True)
        self.new_excel(a).leave_open()
        self.assertEqual(1, a.ole_leave_open_calls)

    # --- the factories declare the bridge cleanup steps --------------------
    #
    # All three declare the identical steps; that is correct even though
    # connect's own record is (almost) never auto-created -- the steps are a
    # property of the instance, and only the bridge knows whether they
    # should ever run.

    def test_create_declares_displayalerts_then_quit_cleanup_steps(self):
        client = FakeClient(FakeComApplication(created=True))
        Excel.create(client=client)
        self.assertEqual([Excel.CLEANUP_STEPS], client.create_cleanups)

    def test_connect_declares_the_same_cleanup_steps(self):
        client = FakeClient(FakeComApplication(created=False))
        Excel.connect(client=client)
        self.assertEqual([Excel.CLEANUP_STEPS], client.connect_cleanups)

    def test_connect_or_create_declares_the_same_cleanup_steps(self):
        client = FakeClient(FakeComApplication(created=True))
        Excel.connect_or_create(client=client)
        self.assertEqual([Excel.CLEANUP_STEPS], client.connect_or_create_cleanups)

    def test_cleanup_steps_are_displayalerts_off_then_quit(self):
        self.assertEqual({'steps': [['DisplayAlerts=', False], ['Quit']]}, Excel.CLEANUP_STEPS)

    # --- the factories use the package default client ----------------------

    def test_create_uses_the_modules_default_client_when_none_given(self):
        fake_app = FakeComApplication(created=True)
        fake_client = FakeClient(fake_app)
        original = wineole.default_client
        wineole.default_client = lambda: fake_client
        try:
            xl = Excel.create()
            self.assertIs(fake_app, xl.ole)
            self.assertEqual(['Excel.Application'], fake_client.create_calls)
        finally:
            wineole.default_client = original

    # --- resolution table: workbook part ------------------------------------

    def test_workbook_active_form_uses_active_workbook(self):
        wb = FakeComWorkbook('Book1', worksheets=[FakeComWorksheet('Sheet1')])
        a = FakeComApplication(created=True, workbooks=[wb], active_workbook=wb)
        result = self.new_excel(a)['[]']
        self.assertIsInstance(result, Book)
        self.assertIs(wb, result.ole)

    def test_workbook_active_form_raises_when_nothing_is_open(self):
        a = FakeComApplication(created=True, active_workbook=None)
        with self.assertRaises(RuntimeError) as ctx:
            self.new_excel(a)['[]']
        self.assertIn('active workbook', str(ctx.exception).lower())

    def test_workbook_new_form_adds_a_workbook(self):
        a = FakeComApplication(created=True, workbooks=[])
        result = self.new_excel(a)['[:new]']
        self.assertIsInstance(result, Book)
        self.assertEqual(1, a.Workbooks().add_calls)
        self.assertIs(a.Workbooks().items[-1], result.ole)

    def test_workbook_named_form_looks_up_by_name(self):
        wb = FakeComWorkbook('Sales', worksheets=[FakeComWorksheet('Sheet1')])
        a = FakeComApplication(created=True, workbooks=[wb])
        result = self.new_excel(a)['[Sales]Sheet1!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1'], wb.worksheets.items[0].range_calls)

    # --- resolution table: worksheet part -----------------------------------

    def test_worksheet_active_form_uses_active_sheet(self):
        active = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, active_sheet=active)
        result = self.new_excel(a)['!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1'], active.range_calls)

    def test_worksheet_active_form_raises_when_nothing_is_open(self):
        a = FakeComApplication(created=True, active_sheet=None)
        with self.assertRaises(RuntimeError) as ctx:
            self.new_excel(a)['!A1']
        self.assertIn('active worksheet', str(ctx.exception).lower())

    def test_worksheet_new_form_adds_after_the_last_sheet(self):
        sheet1 = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, worksheets=[sheet1])
        result = self.new_excel(a)[':new!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual([sheet1], a.Worksheets().add_after_calls)
        self.assertEqual(['A1'], a.Worksheets().items[-1].range_calls)

    def test_worksheet_new_form_without_a_range_returns_the_sheet(self):
        sheet1 = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, worksheets=[sheet1])
        result = self.new_excel(a)[':new!']
        self.assertIsInstance(result, Sheet)
        self.assertEqual([sheet1], a.Worksheets().add_after_calls)

    def test_worksheet_first_form(self):
        s1, s2 = FakeComWorksheet('Sheet1'), FakeComWorksheet('Sheet2')
        a = FakeComApplication(created=True, worksheets=[s1, s2])
        result = self.new_excel(a)[':first!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1'], s1.range_calls)
        self.assertEqual([], s2.range_calls)

    def test_worksheet_last_form(self):
        s1, s2 = FakeComWorksheet('Sheet1'), FakeComWorksheet('Sheet2')
        a = FakeComApplication(created=True, worksheets=[s1, s2])
        result = self.new_excel(a)[':last!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1'], s2.range_calls)
        self.assertEqual([], s1.range_calls)

    def test_worksheet_digit_index_form(self):
        s1, s2 = FakeComWorksheet('Sheet1'), FakeComWorksheet('Sheet2')
        a = FakeComApplication(created=True, worksheets=[s1, s2])
        result = self.new_excel(a)['2!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1'], s2.range_calls)

    def test_worksheet_named_form(self):
        s1 = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, worksheets=[s1])
        result = self.new_excel(a)['Sheet1!A1']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1'], s1.range_calls)

    # --- resolution table: bare range and two-integer forms -----------------

    def test_bare_range_uses_the_active_sheet(self):
        active = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, active_sheet=active)
        result = self.new_excel(a)['A1:B2']
        self.assertIsInstance(result, Range)
        self.assertEqual(['A1:B2'], active.range_calls)

    def test_two_integer_form_uses_the_active_sheet(self):
        active = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, active_sheet=active)
        result = self.new_excel(a)[2, 3]
        self.assertIsInstance(result, Range)
        self.assertEqual([[2, 3]], active.cells_calls)

    def test_two_integer_form_raises_when_nothing_is_open(self):
        a = FakeComApplication(created=True, active_sheet=None)
        with self.assertRaises(RuntimeError) as ctx:
            self.new_excel(a)[2, 3]
        self.assertIn('active worksheet', str(ctx.exception).lower())

    # --- resolution table: raw-name fallback --------------------------------

    def test_raw_name_fallback_wraps_worksheets_item(self):
        s1 = FakeComWorksheet('My Report')
        a = FakeComApplication(created=True, worksheets=[s1])
        result = self.new_excel(a)['My Report']
        self.assertIsInstance(result, Sheet)
        self.assertIs(s1, result.ole)

    # --- []= ------------------------------------------------------------

    def test_subscript_assign_writes_through_the_resolved_range(self):
        active = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, active_sheet=active)
        xl = self.new_excel(a)
        xl['A1:B1'] = 7
        self.assertEqual(7, active.ranges[-1].written)

    def test_subscript_assign_with_two_integers(self):
        active = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, active_sheet=active)
        xl = self.new_excel(a)
        xl[2, 3] = 9
        self.assertEqual([[2, 3]], active.cells_calls)
        self.assertEqual(9, active.ranges[-1].written)

    def test_assigning_to_a_worksheet_only_address_raises(self):
        s1 = FakeComWorksheet('Sheet1')
        a = FakeComApplication(created=True, worksheets=[s1])
        xl = self.new_excel(a)
        with self.assertRaises(ValueError) as ctx:
            xl['Sheet1!'] = 0
        self.assertIn('range', str(ctx.exception))

    def test_assigning_to_a_workbook_only_address_raises(self):
        wb = FakeComWorkbook('Book1', worksheets=[FakeComWorksheet('Sheet1')])
        a = FakeComApplication(created=True, workbooks=[wb], active_workbook=wb)
        xl = self.new_excel(a)
        with self.assertRaises(ValueError) as ctx:
            xl['[]'] = 0
        self.assertIn('range', str(ctx.exception))

    def test_an_unsupported_index_raises_a_type_error(self):
        a = FakeComApplication(created=True)
        xl = self.new_excel(a)
        for bad in [1.5, (1, 2, 3), None, (1, 'A')]:
            with self.assertRaises(TypeError, msg=f"{bad!r} must be refused"):
                xl[bad]

    # --- the version reaches every Sheet and Book Excel builds --------------

    def test_version_reaches_sheet_and_book(self):
        wb = FakeComWorkbook('Book1', worksheets=[FakeComWorksheet('Sheet1')])
        a = FakeComApplication(created=True, version='12.0', workbooks=[wb],
                               worksheets=[FakeComWorksheet('Sheet1')], active_workbook=wb)
        xl = self.new_excel(a)

        # XFD1 is beyond Excel 11's IV/65536 grid, but within Excel 12's --
        # so this only succeeds if the captured version (not a hardcoded
        # 11.0) reached the Sheet Excel built for the lookup.
        self.assertIsInstance(xl['Sheet1!XFD1'], Range)

        # And into the Book Excel builds, via the Sheet the Book builds.
        book = xl['[]']
        self.assertIsInstance(book, Book)
        self.assertIsNotNone(book.sheet('Sheet1')['XFD1'])

    # --- no_alert / no_update restore what was there, not a hardcoded True --

    def test_no_alert_restores_a_pre_existing_false(self):
        a = FakeComApplication(created=True, display_alerts=False)
        xl = self.new_excel(a)
        ran = False
        with xl.no_alert():
            ran = True
            self.assertIs(False, a.DisplayAlerts())
        self.assertTrue(ran)
        self.assertIs(False, a.DisplayAlerts(),
                      'no_alert must restore whatever was there before, not hardcode True')

    def test_no_alert_restores_a_pre_existing_true(self):
        a = FakeComApplication(created=True, display_alerts=True)
        xl = self.new_excel(a)
        with xl.no_alert():
            self.assertIs(False, a.DisplayAlerts())
        self.assertIs(True, a.DisplayAlerts())

    def test_no_alert_restores_even_when_the_block_raises(self):
        a = FakeComApplication(created=True, display_alerts=False)
        xl = self.new_excel(a)
        with self.assertRaises(Boom):
            with xl.no_alert():
                raise Boom()
        self.assertIs(False, a.DisplayAlerts())

    # An exception raised inside the finally REPLACES whatever the block
    # raised. Unguarded, a block that failed while the application was also
    # going away reported the cleanup instead of the real error.
    def test_the_blocks_own_exception_survives_a_restore_that_fails(self):
        a = FakeComApplication(created=False)
        a.fail_display_alerts_restore = True
        xl = self.new_excel(a)
        with self.assertRaises(Boom) as ctx:
            with xl.no_alert():
                raise Boom('the real problem')
        self.assertEqual('the real problem', str(ctx.exception),
                         "the restore's own failure must not stand in for what the block raised")

    # And with no exception of its own to defend, a failed restore must
    # still not manufacture one.
    def test_a_failing_restore_alone_does_not_raise(self):
        a = FakeComApplication(created=False)
        a.fail_display_alerts_restore = True
        xl = self.new_excel(a)
        result = []
        with xl.no_alert():
            result.append('fine')
        self.assertEqual(['fine'], result)

    # The application dying DURING the block leaves no_alert's finally with
    # nothing to put DisplayAlerts back on, and that is not an error.
    def test_an_application_that_dies_during_the_block_is_not_an_error(self):
        a = FakeComApplication(created=True)
        xl = self.new_excel(a)
        with xl.no_alert():
            a.Quit()
        self.assertEqual(1, a.quit_calls, 'the Quit itself must still have happened')

    # --- a failed leading read must not write None into the flag ------------
    #
    # Measured on live Excel 11: writing nil to either flag sets it to
    # FALSE, silently -- so a transient failure on the read followed by a
    # restore of None would leave the flag false for the rest of the
    # session. The read is hoisted out of the protected region so a failing
    # read raises straight to the caller with no restore attempted.

    def test_no_alert_when_the_leading_read_raises_the_caller_sees_it_and_nothing_is_written(self):
        a = FakeComApplication(created=True)
        a.fail_display_alerts_read = True
        xl = self.new_excel(a)
        with self.assertRaises(RuntimeError) as ctx:
            with xl.no_alert():
                self.fail('must not run the block')
        self.assertIn('exploded', str(ctx.exception))
        self.assertEqual([], a.display_alerts_history,
                         'a failed read must never reach the setter')

    def test_no_update_when_the_leading_read_raises_the_caller_sees_it_and_nothing_is_written(self):
        a = FakeComApplication(created=True)
        a.fail_screen_updating_read = True
        xl = self.new_excel(a)
        with self.assertRaises(RuntimeError) as ctx:
            with xl.no_update():
                self.fail('must not run the block')
        self.assertIn('exploded', str(ctx.exception))
        self.assertEqual([], a.screen_updating_history,
                         'a failed read must never reach the setter')

    # --- the flag is only touched inside a `with` ---------------------------
    #
    # Ruby needed a guard here because a block-less `no_alert` set the flag
    # and then died on the bare yield. A @contextmanager cannot: the body
    # does not start until __enter__ does. These pin that, so the property
    # is not lost to a later refactor into an eager helper.

    def test_no_alert_without_a_with_statement_touches_nothing(self):
        a = FakeComApplication(created=True)
        xl = self.new_excel(a)
        xl.no_alert()
        self.assertEqual([], a.display_alerts_history)

    def test_no_update_without_a_with_statement_touches_nothing(self):
        a = FakeComApplication(created=True)
        xl = self.new_excel(a)
        xl.no_update()
        self.assertEqual([], a.screen_updating_history)

    def test_no_update_restores_a_pre_existing_false(self):
        a = FakeComApplication(created=True, screen_updating=False)
        xl = self.new_excel(a)
        with xl.no_update():
            self.assertIs(False, a.ScreenUpdating())
        self.assertIs(False, a.ScreenUpdating())

    def test_no_update_restores_a_pre_existing_true(self):
        a = FakeComApplication(created=True, screen_updating=True)
        xl = self.new_excel(a)
        with xl.no_update():
            self.assertIs(False, a.ScreenUpdating())
        self.assertIs(True, a.ScreenUpdating())

    # --- show / hide ------------------------------------------------------

    def test_show_sets_visible_true(self):
        a = FakeComApplication(created=True)
        self.new_excel(a).show()
        self.assertIs(True, a.visible)

    def test_hide_sets_visible_false(self):
        a = FakeComApplication(created=True)
        self.new_excel(a).hide()
        self.assertIs(False, a.visible)

    # --- passthrough ------------------------------------------------------

    def test_ole_exposes_the_underlying_proxy(self):
        a = FakeComApplication(created=True)
        self.assertIs(a, self.new_excel(a).ole)

    def test_unknown_methods_go_to_com(self):
        a = FakeComApplication(created=True, version='11.0')
        self.assertEqual('11.0', self.new_excel(a).Version())

    # --- the wrapper is not pulled in by the core package -------------------

    def test_importing_wineole_does_not_import_the_office_wrapper(self):
        # In a fresh interpreter, so this test file's own imports cannot
        # make the answer come out right by accident.
        out = subprocess.run(
            [sys.executable, '-c',
             "import sys, wineole; print('wineole.msoffice' in sys.modules)"],
            cwd=os.path.join(os.path.dirname(__file__), '..'),
            stdout=subprocess.PIPE, text=True, check=True,
        ).stdout.strip()
        self.assertEqual('False', out,
                         'the core is a general-purpose COM bridge -- someone who wants it '
                         'must not have to carry an Excel wrapper to get it')


if __name__ == '__main__':
    unittest.main()
