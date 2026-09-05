import contextlib
import os
import queue
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.excel_integration import ExcelIntegrationMixin

from wineole.errors import RemoteError
from wineole.msoffice import Controls, Excel


class MSOfficeControlsIntegrationTest(ExcelIntegrationMixin, unittest.TestCase):
    """Pins the controls and forms layer against real Excel, running under
    Wine. The unit files test every piece against fakes; this is where the
    OLEObjects.Add trap, the event path and the UserForm lifecycle are
    exercised for real.

    Each test gets its own Excel, quit in the context manager's own unwind.
    """

    @contextlib.contextmanager
    def wrapped_excel(self):
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                xl.hide()
                with xl.no_alert():
                    xl.ole.Workbooks().Add()
                    yield xl

    def wait_for(self, fired, what):
        """One arrival, or a verdict. A bare get() would turn a broken event
        path into a hung suite with no output."""
        try:
            return fired.get(timeout=30)
        except queue.Empty:
            self.fail(f"{what} did not arrive within 30s")

    # --- placement ---------------------------------------------------------

    # Every kind of all three families, placed and read back by Name. The
    # tables are the claim the README makes about what can be placed at all.
    def test_every_kind_of_all_three_families_places_and_reads_back_its_name(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']

            for index, kind in enumerate(Controls.FORM_KINDS):
                name = 'F%d' % index
                ctl = sheet.form_controls.add(kind, name=name, left=10,
                                              top=10 + index * 22, width=60,
                                              height=18)
                self.assertEqual(name, ctl.ole.Name(), kind)

            for index, kind in enumerate(Controls.MSFORMS_KINDS):
                name = 'X%d' % index
                ctl = sheet.activex.add(kind, name=name, left=120,
                                        top=10 + index * 22, width=60,
                                        height=18)
                self.assertEqual(name, ctl.ole.Name(), kind)

            form = xl['[]'].forms.add('KindsForm')
            try:
                for index, kind in enumerate(Controls.MSFORMS_KINDS):
                    name = 'U%d' % index
                    ctl = form.controls.add(kind, name=name, left=6,
                                            top=6 + index * 20, width=60,
                                            height=18)
                    self.assertEqual(name, ctl.ole.Name(), kind)
            finally:
                form.unload()

    def test_at_takes_the_ranges_own_four_values(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            box = sheet['B2:C4'].ole
            expected = (box.Left(), box.Top(), box.Width(), box.Height())

            ctl = sheet.activex.add('command_button', name='AtBox', at='B2:C4')

            self.assertAlmostEqual(expected[0], ctl.ole.Left(), places=1)
            self.assertAlmostEqual(expected[1], ctl.ole.Top(), places=1)
            self.assertAlmostEqual(expected[2], ctl.ole.Width(), places=1)
            self.assertAlmostEqual(expected[3], ctl.ole.Height(), places=1)

    # The trap the wrapper exists to make unreachable: on Excel 11
    # OLEObjects.Add with Left and Top but no Width/Height fails outright.
    # Both halves matter -- that it still fails, and that the wrapper path
    # does not.
    def test_the_raw_add_trap_still_fails_and_the_wrapper_path_does_not(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            with self.assertRaises(RemoteError) as caught:
                sheet.ole.OLEObjects().Add(ClassType='Forms.CommandButton.1',
                                           Left=10, Top=10)
            self.assertIn('0x800A03EC', str(caught.exception))

            ctl = sheet.activex.add('command_button', name='Safe', at='D2:E3')
            self.assertEqual('Safe', ctl.ole.Name())

    # --- handlers ----------------------------------------------------------

    # Setting Value on an MSForms CommandButton fires Click; that is how a
    # click is triggered without a UI. This is the whole claim of the
    # "COM events reach the client" row of the README's table.
    def test_an_activex_click_reaches_a_python_callback(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            ctl = sheet.activex.add('command_button', name='Clicker', left=10,
                                    top=10, width=90, height=24, caption='Go')
            fired = queue.Queue()
            ctl.on('Click', lambda *args: fired.put('clicked'))

            ctl.ole_object.Value = True

            self.assertEqual('clicked', self.wait_for(fired, 'the Click'))

    def test_a_vba_handler_runs_for_both_an_activex_and_a_form_control(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']

            ok = sheet.activex.add('command_button', name='Okay', left=10,
                                   top=10, width=90, height=24)
            ok.vba('Click', 'Range("A1").Value = 11')
            ok.ole_object.Value = True
            self.assertEqual(11, sheet['A1'].ole.Value(),
                             'the sheet-module handler ran inside Excel')

            go = sheet.form_controls.add('button', name='Go', left=10, top=60,
                                         width=90, height=24)
            go.vba('Range("A2").Value = 22')
            self.assertEqual('Go_Click', go.ole.OnAction(),
                             'a form control is bound by OnAction, not by events')
            xl.ole.Run(go.ole.OnAction())
            self.assertEqual(22, sheet['A2'].ole.Value())

    # --- the UserForm lifecycle -------------------------------------------

    # The second arrival is the load-bearing one: it says the subscription
    # survived a hide and a show. If it does NOT arrive, re-run this test
    # once to confirm, then apply the documented fallback from the Ruby
    # design spec (2026-09-04-excel-controls-design.md, M2): hide closes the
    # runtime Events, callers re-register after show, the README's
    # Constraints list gains that line, and this assertion becomes a
    # re-registration before the second Value set. Do not delete the
    # assertion.
    def test_a_userform_control_click_reaches_python_across_hide_and_show(self):
        with self.wrapped_excel() as xl:
            form = xl['[]'].forms.add('AppForm')
            try:
                ok = form.controls.add('command_button', name='OK', left=10,
                                       top=10, width=80, height=24, caption='OK')
                fired = queue.Queue()
                ok.on('Click', lambda *args: fired.put('clicked'))

                form.show()
                self.assertTrue(form.shown())
                ok.runtime().Value = True
                self.assertEqual('clicked', self.wait_for(fired, 'the first Click'))

                form.hide()
                self.assertFalse(form.shown())
                form.show()
                ok.runtime().Value = True
                self.assertEqual('clicked',
                                 self.wait_for(fired, 'the second Click'))
            finally:
                # Before Excel is quit, and before the connection goes: unload
                # closes the subscriptions on our side rather than leaving
                # them to fail when the objects are gone.
                form.unload()

    # --- the refusals, against the real thing ------------------------------

    def test_a_duplicate_name_is_refused_and_a_refused_property_rolls_back(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            sheet.activex.add('command_button', name='Only', left=10, top=10,
                              width=80, height=24)
            before = sheet.ole.Shapes().Count()

            with self.assertRaises(ValueError):
                sheet.form_controls.add('button', name='Only', left=10, top=40,
                                        width=80, height=24)
            self.assertEqual(before, sheet.ole.Shapes().Count(),
                             'the duplicate was refused before Excel was asked')

            # An Image has no Caption, so Excel refuses the put -- and the
            # control that was just added must be gone again.
            with self.assertRaises(RemoteError):
                sheet.activex.add('image', name='Pic', left=10, top=80,
                                  width=60, height=60, caption='x')
            self.assertEqual(before, sheet.ole.Shapes().Count(),
                             'the refused control was deleted again')


if __name__ == '__main__':
    unittest.main()
