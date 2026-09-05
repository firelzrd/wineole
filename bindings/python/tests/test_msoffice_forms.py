import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import (FakeBookForForms, FakeClient, FakeRuntimeControl,
                              FakeWorkbookForForms, remote_error)

from wineole.errors import RemoteError
from wineole.msoffice.book import Book
from wineole.msoffice.controls import UserFormControls
from wineole.msoffice.forms import Form, Forms

BOX = {'left': 1, 'top': 2, 'width': 3, 'height': 4}

HELPER = ('Function WineOLE_Form_AppForm() As Object\n'
          '    Set WineOLE_Form_AppForm = AppForm\n'
          'End Function\n'
          'Sub WineOLE_Show_AppForm()\n'
          '    AppForm.Show 0\n'
          'End Sub\n'
          'Sub WineOLE_Hide_AppForm()\n'
          '    AppForm.Hide\n'
          'End Sub\n'
          'Sub WineOLE_Unload_AppForm()\n'
          '    Unload AppForm\n'
          'End Sub\n')


class FormsTest(unittest.TestCase):

    def setUp(self):
        self.book = FakeBookForForms()
        self.forms = Forms(self.book)

    def test_add_makes_a_form_component_and_wraps_its_designer(self):
        form = self.forms.add('AppForm')
        self.assertIsInstance(form, Form)
        self.assertEqual('AppForm', form.name)
        self.assertEqual(3, form.component.Type())
        self.assertIs(form.component.Designer(), form.ole)
        self.assertIs(self.book.vba.components['AppForm'], form.component)

    def test_add_checks_the_name_before_touching_vba(self):
        with self.assertRaises(ValueError) as caught:
            self.forms.add('App Form')
        self.assertIn('name:', str(caught.exception))
        self.assertNotIn('App Form', self.book.vba.components)

    def test_add_writes_the_helper_block_into_the_wrapper_module(self):
        self.forms.add('AppForm')
        self.assertEqual([{'code': Forms.helper('AppForm'),
                           'name': 'form_AppForm', 'into': None}],
                         self.book.vba.writes)

    def test_the_helper_text(self):
        self.assertEqual(HELPER, Forms.helper('AppForm'))

    def test_a_taken_name_is_add_components_refusal(self):
        self.forms.add('AppForm')
        with self.assertRaises(ValueError) as caught:
            self.forms.add('AppForm')
        self.assertIn('already has a VBA component named', str(caught.exception))

    def test_a_refused_helper_write_removes_the_component_and_reraises(self):
        error = remote_error('0x800A03EC')
        self.book.vba.fail_write_with = error

        with self.assertRaises(RemoteError) as caught:
            self.forms.add('AppForm')

        self.assertIs(error, caught.exception, 'the original error, unchanged')
        self.assertNotIn('AppForm', self.book.vba.components)
        self.assertEqual(['AppForm'], self.book.vba.removed)

    def test_lookup_rebinds_a_userform_and_nothing_else(self):
        self.forms.add('AppForm')
        found = self.forms['AppForm']
        self.assertIsInstance(found, Form)
        self.assertIs(self.book.vba.components['AppForm'], found.component)
        self.assertIsNone(self.forms['Module1'], 'a standard module is not a form')
        self.assertIsNone(self.forms['Nope'])

    def test_lookup_rewrites_the_helper_so_a_reopened_workbook_can_show_the_form(self):
        self.forms.add('AppForm')
        self.book.vba.writes.clear()
        self.forms['AppForm']
        self.assertEqual(['form_AppForm'],
                         [write['name'] for write in self.book.vba.writes])

    # Only the lookup answers "not there" with None. A refused helper write
    # is a real failure -- a locked project, a refused AddFromString -- and
    # swallowing it would hand back a None for a form that is right there.
    def test_lookup_raises_when_the_helper_write_is_refused(self):
        self.forms.add('AppForm')
        error = remote_error('0x800A03EC')
        self.book.vba.fail_write_with = error

        with self.assertRaises(RemoteError) as caught:
            self.forms['AppForm']

        self.assertIs(error, caught.exception)
        self.assertIsNone(self.forms['Nope'], 'a missing name is still None')

    def test_book_forms_is_memoized(self):
        book = Book(FakeWorkbookForForms(), FakeClient(), '11.0')
        self.assertIsInstance(book.forms, Forms)
        self.assertIs(book.forms, book.forms)


class FormTest(unittest.TestCase):

    def setUp(self):
        self.book = FakeBookForForms()
        self.form = Forms(self.book).add('AppForm')
        self.app = self.book.ole.Application()

    def test_show_runs_the_show_helper_qualified_with_the_workbook_name(self):
        self.assertIs(self.form, self.form.show())
        self.assertEqual(["'Book1'!WineOLE_Show_AppForm"], self.app.runs)

    def test_instance_is_fetched_once_through_the_form_function(self):
        first = self.form.instance()
        self.assertIs(first, self.form.instance())
        self.assertEqual(["'Book1'!WineOLE_Form_AppForm"], self.app.runs)
        self.assertIs(self.app.instances[0], first)

    def test_hide_runs_the_hide_helper_and_shown_reads_the_instance(self):
        self.form.instance().mark_shown()
        self.assertTrue(self.form.shown())
        self.assertIs(self.form, self.form.hide())
        self.assertEqual(["'Book1'!WineOLE_Form_AppForm",
                          "'Book1'!WineOLE_Hide_AppForm"], self.app.runs)
        self.assertEqual([], self.form.instance().calls,
                         'hide must not call the instance: its Hide is not '
                         'dispatchable out of process')

    def test_runtime_controls_are_cached_per_name(self):
        first = self.form.runtime_control('OK')
        self.assertIs(first, self.form.runtime_control('OK'))
        self.assertIsNot(first, self.form.runtime_control('Cancel'))
        self.assertIsInstance(first, FakeRuntimeControl)

    def test_unload_closes_every_events_object_before_running_unload_and_forgets_the_instance(self):
        instance = self.form.instance()
        ok = self.form.runtime_control('OK')
        ok.ole_events.on('Click', lambda *args: None)

        self.assertIs(self.form, self.form.unload())

        self.assertEqual(1, ok.ole_events.closed)
        self.assertEqual(1, instance.ole_events.closed)
        self.assertEqual(["'Book1'!WineOLE_Form_AppForm",
                          "'Book1'!WineOLE_Unload_AppForm"], self.app.runs)
        self.assertIsNot(instance, self.form.instance(),
                         'the next instance is fetched afresh')
        self.assertIsNot(ok, self.form.runtime_control('OK'))

    def test_unload_without_an_instance_just_runs_the_helper(self):
        self.form.unload()
        self.assertEqual(["'Book1'!WineOLE_Unload_AppForm"], self.app.runs)

    def test_unknown_methods_reach_the_designer(self):
        self.assertEqual(240, self.form.Width())


class UserFormControlsTest(unittest.TestCase):

    def setUp(self):
        self.book = FakeBookForForms()
        self.form = Forms(self.book).add('AppForm')
        self.designer_controls = self.form.ole.Controls()

    def test_controls_is_memoized(self):
        self.assertIs(self.form.controls, self.form.controls)
        self.assertIsInstance(self.form.controls, UserFormControls)

    def test_add_names_the_control_at_add_time_then_sizes_it(self):
        ctl = self.form.controls.add('command_button', name='OK', **BOX)
        self.assertEqual([['Forms.CommandButton.1', 'OK']],
                         self.designer_controls.add_calls)
        self.assertEqual([['Left', 1], ['Top', 2], ['Width', 3], ['Height', 4]],
                         ctl.ole.puts_seen)
        self.assertEqual('userform', ctl.family)
        self.assertEqual('command_button', ctl.kind)
        self.assertIs(ctl.ole, ctl.ole_object)

    def test_properties_follow_the_box(self):
        ctl = self.form.controls.add('command_button', name='OK', caption='OK',
                                     **BOX)
        self.assertEqual(['Caption', 'OK'], ctl.ole.puts_seen[-1])

    def test_at_is_refused_on_a_form(self):
        with self.assertRaises(ValueError) as caught:
            self.form.controls.add('command_button', name='OK', at='B2')
        self.assertIn('a UserForm has no cells', str(caught.exception))
        self.assertEqual([], self.designer_controls.add_calls)

    def test_a_refused_property_removes_the_control_and_reraises(self):
        with self.assertRaises(RemoteError):
            self.form.controls.add('command_button', name='OK', bogus=1, **BOX)
        self.assertEqual(1, len(self.designer_controls.add_calls))
        self.assertEqual(0, self.designer_controls.Count())

    def test_a_duplicate_name_is_refused_before_add(self):
        self.form.controls.add('command_button', name='OK', **BOX)
        with self.assertRaises(ValueError) as caught:
            self.form.controls.add('label', name='OK', **BOX)
        self.assertIn("this UserForm already has a control named 'OK'",
                      str(caught.exception))
        self.assertEqual(1, len(self.designer_controls.add_calls))

    def test_lookup_rebinds_with_no_kind(self):
        self.form.controls.add('text_box', name='Name', **BOX)
        found = self.form.controls['Name']
        self.assertEqual('userform', found.family)
        self.assertIsNone(found.kind,
                          'a placed MSForms control reports no ProgID')
        self.assertIs(self.designer_controls.Item('Name'), found.ole)
        self.assertIsNone(self.form.controls['Nope'])

    def test_a_callback_handler_lands_on_the_runtime_control(self):
        ctl = self.form.controls.add('command_button', name='OK', **BOX)
        ctl.on('Click', lambda *args: None)
        self.assertEqual(
            1, len(self.form.runtime_control('OK').ole_events.on_calls))

    def test_a_vba_handler_goes_into_the_form_module(self):
        self.form.controls.add('command_button', name='OK',
                               **BOX).vba('Click', 'Beep')
        self.assertEqual({'code': 'Private Sub OK_Click()\n    Beep\nEnd Sub',
                          'name': 'OK_Click', 'into': 'AppForm'},
                         self.book.vba.writes[-1])


if __name__ == '__main__':
    unittest.main()
