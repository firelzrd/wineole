import contextlib
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import (_Recorder, FakeComSheetForControls,
                              FakeComSheetForGeometry, FakeControlComObject,
                              FakeFormForControl, FakeItemHost, FakeVBAWriter)

from wineole.errors import RemoteError
from wineole.msoffice.controls import Control, Controls
from wineole.msoffice.sheet import Sheet
from wineole.msoffice.vba import VBA, VBAAccessDenied


def geometry_sheet():
    """A real Sheet around a fake COM worksheet: geometry() reads its box
    through `sheet[at].ole`, so the Sheet wrapper's own addressing rules --
    including its refusal of another sheet's address -- are part of what is
    being tested."""
    return Sheet(FakeComSheetForGeometry(), '11.0')


class ControlsModuleTest(unittest.TestCase):

    # --- kind tables ---------------------------------------------------

    def test_form_kinds_map_to_the_nine_legacy_collections(self):
        self.assertEqual(
            ['Buttons', 'CheckBoxes', 'OptionButtons', 'ListBoxes', 'DropDowns',
             'Spinners', 'ScrollBars', 'Labels', 'GroupBoxes'],
            list(Controls.FORM_KINDS.values()))
        self.assertEqual('CheckBoxes', Controls.form_collection_for('check_box'))

    def test_an_unknown_form_kind_lists_the_valid_ones(self):
        with self.assertRaises(ValueError) as caught:
            Controls.form_collection_for('command_button')
        message = str(caught.exception)
        self.assertIn("unknown form control kind 'command_button'", message)
        self.assertIn("'button', 'check_box', 'option_button'", message)
        self.assertIn('sheet.activex', message,
                      'it must point at the family that has this kind')

    # Ruby has a separate branch for a String kind ("form controls have no
    # ProgID"); in Python a ProgID and a kind are both str, so a ProgID lands
    # in the unknown-kind message -- which already says where to take it.
    def test_a_progid_is_not_a_form_control_kind(self):
        with self.assertRaises(ValueError) as caught:
            Controls.form_collection_for('Forms.CommandButton.1')
        message = str(caught.exception)
        self.assertIn("unknown form control kind 'Forms.CommandButton.1'", message)
        self.assertIn('sheet.activex', message)

    def test_msforms_kinds_map_to_progids(self):
        self.assertEqual('Forms.CommandButton.1',
                         Controls.progid_for('command_button'))
        self.assertEqual('Forms.Image.1', Controls.progid_for('image'))
        self.assertEqual(11, len(Controls.MSFORMS_KINDS))

    def test_a_string_kind_is_the_progid_verbatim(self):
        self.assertEqual('MSComctlLib.ProgCtrl.2',
                         Controls.progid_for('MSComctlLib.ProgCtrl.2'))

    def test_a_kind_that_is_not_a_str_is_a_type_error(self):
        for bad in (None, 1, ['button']):
            with self.assertRaises(TypeError, msg=repr(bad)) as caught:
                Controls.progid_for(bad)
            self.assertIn('a control kind must be a str', str(caught.exception))
            with self.assertRaises(TypeError, msg=repr(bad)):
                Controls.form_collection_for(bad)

    def test_kind_is_recovered_from_a_progid(self):
        self.assertEqual('toggle_button',
                         Controls.kind_for_progid('Forms.ToggleButton.1'))
        self.assertEqual('Other.Ctrl.1', Controls.kind_for_progid('Other.Ctrl.1'))

    def test_form_control_types_skip_the_edit_box(self):
        self.assertEqual('button', Controls.FORM_CONTROL_TYPES[0])
        self.assertEqual('spinner', Controls.FORM_CONTROL_TYPES[9])
        self.assertIsNone(Controls.FORM_CONTROL_TYPES.get(3),
                          'xlEditBox exists only on dialog sheets')

    # --- names ---------------------------------------------------------

    def test_a_vba_identifier_is_a_valid_name(self):
        self.assertIsNone(Controls.check_name('OK_Button2'))
        self.assertIsNone(Controls.check_name('a' * 31))

    def test_names_that_cannot_appear_in_name_click_are_refused(self):
        for bad in ('', '1st', 'my button', 'ok-button', 'a' * 32):
            with self.assertRaises(ValueError, msg=repr(bad)) as caught:
                Controls.check_name(bad)
            self.assertIn('name:', str(caught.exception))
        for bad in (None, 1):
            with self.assertRaises(TypeError, msg=repr(bad)) as caught:
                Controls.check_name(bad)
            self.assertIn('a control name must be a str', str(caught.exception))

    def test_event_names_are_vba_identifiers(self):
        self.assertIsNone(Controls.check_event('KeyDown'))
        with self.assertRaises(ValueError) as caught:
            Controls.check_event('Key Down')
        self.assertIn('event', str(caught.exception))
        with self.assertRaises(TypeError) as caught:
            Controls.check_event(None)
        self.assertIn('an event name must be a str', str(caught.exception))

    def test_check_free_passes_when_the_lookup_raises(self):
        self.assertIsNone(
            Controls.check_free(FakeItemHost('Other'), 'Fresh', 'sheet'))

    def test_check_free_refuses_a_name_the_host_already_has(self):
        with self.assertRaises(ValueError) as caught:
            Controls.check_free(FakeItemHost('Taken'), 'Taken', 'sheet')
        self.assertIn("this sheet already has a control named 'Taken'",
                      str(caught.exception))

    # --- properties ----------------------------------------------------

    def test_snake_case_keys_become_pascal_case(self):
        self.assertEqual('Caption', Controls.pascal('caption'))
        self.assertEqual('BackColor', Controls.pascal('back_color'))
        self.assertEqual('ListFillRange', Controls.pascal('list_fill_range'))
        self.assertEqual('OnAction', Controls.pascal('on_action'))

    def test_put_sends_one_property_assignment(self):
        target = _Recorder()
        Controls.put(target, 'back_color', 255)
        self.assertEqual({'BackColor': 255}, target.writes)

    def test_the_five_host_props(self):
        self.assertEqual(('linked_cell', 'list_fill_range', 'visible',
                          'print_object', 'placement'), Controls.HOST_PROPS)

    # --- geometry ------------------------------------------------------

    def test_at_resolves_through_the_sheet_to_the_ranges_four_values(self):
        sheet = geometry_sheet()
        box = Controls.geometry(sheet=sheet, at='B2:C4', left=None, top=None,
                                width=None, height=None)
        self.assertEqual((10.5, 20.0, 30.0, 40.0), box)
        self.assertEqual(['B2:C4'], sheet.ole.range_calls)

    def test_at_inherits_the_sheets_own_sheet_rule(self):
        sheet = geometry_sheet()
        with self.assertRaises(ValueError) as caught:
            Controls.geometry(sheet=sheet, at='Other!B2', left=None, top=None,
                              width=None, height=None)
        self.assertIn('Other', str(caught.exception))

    def test_four_points_pass_through_unchanged(self):
        box = Controls.geometry(sheet=geometry_sheet(), at=None, left=1, top=2,
                                width=3, height=4)
        self.assertEqual((1, 2, 3, 4), box)

    def test_mixing_at_and_points_raises(self):
        with self.assertRaises(ValueError) as caught:
            Controls.geometry(sheet=geometry_sheet(), at='B2', left=1, top=None,
                              width=None, height=None)
        self.assertIn('not both', str(caught.exception))

    def test_no_position_at_all_raises(self):
        with self.assertRaises(ValueError) as caught:
            Controls.geometry(sheet=geometry_sheet(), at=None, left=None,
                              top=None, width=None, height=None)
        self.assertIn('no position', str(caught.exception))

    def test_a_partial_box_names_what_is_missing(self):
        with self.assertRaises(ValueError) as caught:
            Controls.geometry(sheet=geometry_sheet(), at=None, left=1, top=2,
                              width=None, height=None)
        self.assertIn("missing ['width', 'height']", str(caught.exception))

    def test_a_userform_has_no_cells(self):
        with self.assertRaises(ValueError) as caught:
            Controls.geometry(sheet=None, at='B2', left=None, top=None,
                              width=None, height=None)
        self.assertIn('a UserForm has no cells', str(caught.exception))

    def test_geometry_never_touches_excel_when_it_refuses(self):
        sheet = geometry_sheet()
        with self.assertRaises(ValueError):
            Controls.geometry(sheet=sheet, at='B2', left=1, top=2, width=3,
                              height=4)
        self.assertEqual([], sheet.ole.range_calls)


class ControlTest(unittest.TestCase):

    def form_control(self, writer=None):
        ole = FakeControlComObject()
        return Control(name='Go', kind='button', family='form_control', ole=ole,
                       ole_object=ole, vba=writer or FakeVBAWriter())

    def activex(self, writer=None):
        # The one family whose two handles are genuinely different objects.
        return Control(name='Go', kind='command_button', family='activex',
                       ole=FakeControlComObject(),
                       ole_object=FakeControlComObject(),
                       vba=writer or FakeVBAWriter())

    def userform(self, writer=None, form=None):
        ole = FakeControlComObject()
        return Control(name='OK', kind='command_button', family='userform',
                       ole=ole, ole_object=ole, vba=writer or FakeVBAWriter(),
                       form=form or FakeFormForControl('AppForm'))

    # --- readers and passthrough ---------------------------------------

    def test_readers(self):
        ctl = self.activex()
        self.assertEqual('Go', ctl.name)
        self.assertEqual('command_button', ctl.kind)
        self.assertEqual('activex', ctl.family)
        self.assertIsNot(ctl.ole, ctl.ole_object,
                         'a worksheet ActiveX control is two objects')

    def test_unknown_methods_go_to_ole_object_not_ole(self):
        ctl = self.activex()
        ctl.Caption = 'Go'
        self.assertEqual([['Caption', 'Go']], ctl.ole_object.puts_seen)
        self.assertEqual([], ctl.ole.puts_seen)

    def test_ole_object_is_ole_for_the_other_two_families(self):
        plain = self.form_control()
        self.assertIs(plain.ole, plain.ole_object)
        ctl = self.userform()
        self.assertIs(ctl.ole, ctl.ole_object)

    def test_the_com_member_Object_is_still_reachable(self):
        # The reader is `ole_object`, not `object`, precisely so a bare
        # `ctl.Object()` passes through to the COM member of that name.
        self.assertEqual('the-com-member', self.activex().Object())
        self.assertEqual('the-com-member', self.userform().Object())

    # --- callback handlers -----------------------------------------------

    def test_a_form_control_refuses_a_callback(self):
        with self.assertRaises(ValueError) as caught:
            self.form_control().on('Click', lambda *args: None)
        message = str(caught.exception)
        self.assertIn('form controls have no COM events', message)
        self.assertIn('sheet.activex', message)
        self.assertIsNone(self.form_control().events(),
                          'events() is silently None for this family; only on/off raise')

    def test_off_on_a_form_control_raises_the_same_way(self):
        with self.assertRaises(ValueError):
            self.form_control().off('Click')

    def test_activex_on_reaches_the_inner_objects_events(self):
        ctl = self.activex()

        def callback(*args):
            return None

        self.assertEqual('subscription', ctl.on('Click', callback, args=False))
        self.assertEqual([['Click', callback, False]],
                         ctl.ole_object.ole_events.on_calls)
        self.assertIs(ctl.ole_object.ole_events, ctl.events())

    def test_activex_off_forwards(self):
        ctl = self.activex()
        ctl.off('subscription')
        self.assertEqual(['subscription'], ctl.ole_object.ole_events.off_calls)

    def test_userform_on_reaches_the_runtime_control_not_the_designer(self):
        form = FakeFormForControl('AppForm')
        ctl = self.userform(form=form)
        ctl.on('Click', lambda *args: None)
        self.assertEqual(['OK'], form.runtime_requests)
        self.assertEqual(1, len(form.runtime_control('OK').ole_events.on_calls))
        self.assertEqual([], ctl.ole.ole_events.on_calls,
                         'the design-time control never fires events')

    def test_runtime_is_userform_only(self):
        form = FakeFormForControl('AppForm')
        self.assertIs(form.runtime_control('OK'),
                      self.userform(form=form).runtime())
        with self.assertRaises(ValueError) as caught:
            self.activex().runtime()
        self.assertIn('only a UserForm control has a runtime instance',
                      str(caught.exception))

    # --- VBA handlers --------------------------------------------------

    def test_a_form_control_gets_a_click_sub_and_on_action(self):
        writer = FakeVBAWriter()
        ctl = self.form_control(writer=writer)
        self.assertIs(ctl, ctl.vba('Range("A1").Value = 1\nBeep'))
        self.assertEqual(
            [{'code': 'Sub Go_Click()\n    Range("A1").Value = 1\n    Beep\n'
                      'End Sub',
              'name': 'Go_Click', 'into': None}],
            writer.writes)
        self.assertEqual([['OnAction', 'Go_Click']], ctl.ole.puts_seen)

    def test_a_form_control_takes_no_event_name(self):
        with self.assertRaises(ValueError) as caught:
            self.form_control().vba('Click', 'Beep')
        self.assertIn('fires only Click', str(caught.exception))

    def test_an_activex_handler_is_a_private_sub_in_the_sheet_writer(self):
        writer = FakeVBAWriter()
        self.activex(writer=writer).vba('Click', 'Beep')
        self.assertEqual([{'code': 'Private Sub Go_Click()\n    Beep\nEnd Sub',
                           'name': 'Go_Click', 'into': None}], writer.writes)

    def test_params_go_into_the_signature_verbatim(self):
        writer = FakeVBAWriter()
        self.activex(writer=writer).vba(
            'KeyDown', 'Beep',
            params='ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer')
        self.assertEqual(
            'Private Sub Go_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, '
            'ByVal Shift As Integer)\n    Beep\nEnd Sub',
            writer.writes[0]['code'])
        self.assertEqual('Go_KeyDown', writer.writes[0]['name'])

    def test_a_userform_handler_goes_into_the_form_module(self):
        writer = FakeVBAWriter()
        self.userform(writer=writer).vba('Click', 'Beep')
        self.assertEqual([{'code': 'Private Sub OK_Click()\n    Beep\nEnd Sub',
                           'name': 'OK_Click', 'into': 'AppForm'}], writer.writes)

    def test_activex_and_userform_need_an_event_name(self):
        with self.assertRaises(ValueError) as caught:
            self.activex().vba('Beep')
        self.assertIn('vba(event, body)', str(caught.exception))
        with self.assertRaises(ValueError):
            self.userform().vba('Beep')

    def test_an_event_name_that_is_not_an_identifier_is_refused_before_writing(self):
        writer = FakeVBAWriter()
        with self.assertRaises(ValueError):
            self.activex(writer=writer).vba('Key Down', 'Beep')
        self.assertEqual([], writer.writes)

    def test_blank_lines_in_the_body_stay_blank(self):
        writer = FakeVBAWriter()
        self.activex(writer=writer).vba('Click', 'A\n\nB\n')
        self.assertEqual('Private Sub Go_Click()\n    A\n\n    B\nEnd Sub',
                         writer.writes[0]['code'])


BOX = {'left': 1, 'top': 2, 'width': 3, 'height': 4}

# VBA.denied() consults the real registry to choose its wording, which would
# shell out to `wine reg` from a unit test and make the result depend on the
# host. This is the same seam tests/test_msoffice_vba_api.py replaces.
ENABLED = ('    AccessVBOM    REG_DWORD    0x1\r\n', True)


@contextlib.contextmanager
def with_reg(result):
    original = VBA.__dict__['run_reg']
    VBA.run_reg = staticmethod(lambda args: result)
    VBA.forget_codepage()
    try:
        yield
    finally:
        VBA.run_reg = original
        VBA.forget_codepage()


def controls_sheet(denied=False):
    return Sheet(FakeComSheetForControls(denied=denied), '11.0')


class FormControlsTest(unittest.TestCase):

    def test_readers_are_memoized_and_distinct(self):
        sheet = controls_sheet()
        self.assertIs(sheet.form_controls, sheet.form_controls)
        self.assertIs(sheet.activex, sheet.activex)
        self.assertIsNot(sheet.form_controls, sheet.activex)

    def test_add_places_through_the_kinds_collection_at_the_given_box(self):
        sheet = controls_sheet()
        ctl = sheet.form_controls.add('button', name='Go', **BOX)
        self.assertEqual('form_control', ctl.family)
        self.assertEqual('button', ctl.kind)
        self.assertEqual('Go', ctl.name)
        self.assertEqual([['Buttons', 1, 2, 3, 4]], sheet.ole.add_calls)
        self.assertEqual(['Go'], list(sheet.ole.registry))
        self.assertIs(ctl.ole, ctl.ole_object)

    def test_add_with_at_reads_the_ranges_box(self):
        sheet = controls_sheet()
        ctl = sheet.form_controls.add('check_box', name='Agree', at='B2')
        self.assertEqual(['B2'], sheet.ole.range_calls)
        self.assertEqual([10.5, 20.0, 30.0, 40.0], ctl.ole.box)
        self.assertEqual('CheckBoxes', ctl.ole.collection)

    def test_rename_comes_before_the_properties(self):
        ctl = controls_sheet().form_controls.add('button', name='Go',
                                                 caption='Run', **BOX)
        self.assertEqual([['Name', 'Go'], ['Caption', 'Run']], ctl.ole.puts_seen)

    def test_a_refused_property_deletes_the_control_and_reraises(self):
        sheet = controls_sheet()
        with self.assertRaises(RemoteError) as caught:
            sheet.form_controls.add('button', name='Go', bogus=1, **BOX)
        self.assertIn('0x80020006', str(caught.exception))
        self.assertEqual(1, len(sheet.ole.add_calls), 'the control was added')
        self.assertEqual(0, sheet.ole.Shapes().Count(), 'and then deleted')

    def test_a_refused_rename_deletes_the_control_and_reraises(self):
        sheet = controls_sheet()
        with self.assertRaises(RemoteError) as caught:
            sheet.form_controls.add('button', name='Refused',
                                    caption='never set', **BOX)
        self.assertIn('0x800A03EC', str(caught.exception))
        self.assertEqual(1, len(sheet.ole.add_calls), 'the control was added')
        self.assertEqual(0, sheet.ole.Shapes().Count(), 'and then deleted')

    def test_a_duplicate_name_is_refused_before_excel_is_asked_again(self):
        sheet = controls_sheet()
        sheet.form_controls.add('button', name='Go', **BOX)
        with self.assertRaises(ValueError) as caught:
            sheet.form_controls.add('label', name='Go', **BOX)
        self.assertIn("already has a control named 'Go'", str(caught.exception))
        self.assertEqual(1, len(sheet.ole.add_calls))

    def test_argument_checks_run_before_excel_is_touched(self):
        sheet = controls_sheet()
        with self.assertRaises(ValueError):
            sheet.form_controls.add('command_button', name='Go', **BOX)
        with self.assertRaises(ValueError):
            sheet.form_controls.add('button', name='bad name', **BOX)
        with self.assertRaises(ValueError):
            sheet.form_controls.add('button', name='Go', left=1)
        with self.assertRaises(ValueError):
            sheet.form_controls.add('button', name='Go', at='B2', left=1)
        self.assertEqual([], sheet.ole.add_calls)
        self.assertEqual([], sheet.ole.range_calls)

    def test_lookup_rebinds_by_form_control_type(self):
        sheet = controls_sheet()
        added = sheet.form_controls.add('spinner', name='Spin', **BOX)
        found = sheet.form_controls['Spin']
        self.assertEqual('spinner', found.kind)
        self.assertEqual('form_control', found.family)
        self.assertIs(added.ole, found.ole)

    def test_lookup_is_none_for_a_missing_name_and_for_an_activex_shape(self):
        sheet = controls_sheet()
        self.assertIsNone(sheet.form_controls['Nope'])
        sheet.activex.add('command_button', name='X', **BOX)
        self.assertIsNone(sheet.form_controls['X'])

    def test_a_vba_handler_lands_in_the_wrapper_module_of_the_parent_workbook(self):
        sheet = controls_sheet()
        ctl = sheet.form_controls.add('button', name='Go', **BOX)
        ctl.vba('Beep')
        components = sheet.ole.project.components
        self.assertIn('WineOLE', components)
        self.assertIn('Sub Go_Click()', components['WineOLE'].CodeModule().text)
        self.assertEqual('', components['Sheet1'].CodeModule().text,
                         'not the sheet module')
        self.assertIn(['OnAction', 'Go_Click'], ctl.ole.puts_seen)

    def test_a_denied_project_is_access_denied_not_a_remote_error(self):
        ctl = controls_sheet(denied=True).form_controls.add('button', name='Go',
                                                            **BOX)
        with with_reg(ENABLED):
            with self.assertRaises(VBAAccessDenied):
                ctl.vba('Beep')


class ActiveXControlsTest(unittest.TestCase):

    def test_add_passes_classtype_and_all_four_points_by_name(self):
        sheet = controls_sheet()
        ctl = sheet.activex.add('command_button', name='Go', **BOX)
        self.assertEqual([{'ClassType': 'Forms.CommandButton.1', 'Left': 1,
                           'Top': 2, 'Width': 3, 'Height': 4}],
                         sheet.ole.ole_add_calls)
        self.assertEqual('activex', ctl.family)
        self.assertEqual('command_button', ctl.kind)
        self.assertIs(ctl.ole.inner, ctl.ole_object)
        self.assertEqual(['Go'], list(sheet.ole.registry))

    def test_a_string_kind_is_the_progid_and_stays_the_kind(self):
        sheet = controls_sheet()
        ctl = sheet.activex.add('Other.Ctrl.1', name='Ext', **BOX)
        self.assertEqual('Other.Ctrl.1', sheet.ole.ole_add_calls[0]['ClassType'])
        self.assertEqual('Other.Ctrl.1', ctl.kind)

    def test_properties_are_routed_to_the_host_or_the_inner_object(self):
        ctl = controls_sheet().activex.add('command_button', name='Go',
                                           caption='Go', linked_cell='A1',
                                           back_color=255, **BOX)
        self.assertEqual([['Name', 'Go'], ['LinkedCell', 'A1']],
                         ctl.ole.puts_seen)
        self.assertEqual([['Caption', 'Go'], ['BackColor', 255]],
                         ctl.ole_object.puts_seen)

    def test_a_refused_inner_property_deletes_the_control_and_reraises(self):
        sheet = controls_sheet()
        with self.assertRaises(RemoteError) as caught:
            sheet.activex.add('image', name='Pic', caption='x', **BOX)
        self.assertIn('0x80020006', str(caught.exception))
        self.assertEqual(1, len(sheet.ole.ole_add_calls))
        self.assertEqual(0, sheet.ole.Shapes().Count())

    def test_a_refused_rename_deletes_the_control_and_reraises(self):
        sheet = controls_sheet()
        with self.assertRaises(RemoteError) as caught:
            sheet.activex.add('command_button', name='Refused', **BOX)
        self.assertIn('0x800A03EC', str(caught.exception))
        self.assertEqual(1, len(sheet.ole.ole_add_calls))
        self.assertEqual(0, sheet.ole.Shapes().Count())

    def test_a_name_taken_by_a_form_control_is_refused(self):
        sheet = controls_sheet()
        sheet.form_controls.add('button', name='Twice', **BOX)
        with self.assertRaises(ValueError) as caught:
            sheet.activex.add('command_button', name='Twice', **BOX)
        self.assertIn("already has a control named 'Twice'",
                      str(caught.exception))
        self.assertEqual([], sheet.ole.ole_add_calls)

    def test_argument_checks_run_before_excel_is_touched(self):
        sheet = controls_sheet()
        # Ruby refuses `:button` here as an unknown ActiveX kind; in Python
        # every str is a legal ProgID, so the only kind that can be refused
        # is one that is not a str at all.
        with self.assertRaises(TypeError):
            sheet.activex.add(None, name='Go', **BOX)
        with self.assertRaises(ValueError):
            sheet.activex.add('command_button', name='1Go', **BOX)
        with self.assertRaises(ValueError):
            sheet.activex.add('command_button', name='Go', width=3)
        self.assertEqual([], sheet.ole.ole_add_calls)

    def test_lookup_rebinds_and_recovers_the_kind_from_the_progid(self):
        sheet = controls_sheet()
        sheet.activex.add('toggle_button', name='T', **BOX)
        found = sheet.activex['T']
        self.assertEqual('toggle_button', found.kind)
        self.assertEqual('activex', found.family)
        self.assertIsNone(sheet.activex['nope'])

    def test_lookup_is_none_for_a_form_control_shape(self):
        sheet = controls_sheet()
        sheet.form_controls.add('button', name='Plain', **BOX)
        self.assertIsNone(sheet.activex['Plain'])

    def test_a_vba_handler_lands_in_the_sheets_own_module(self):
        sheet = controls_sheet()
        sheet.activex.add('command_button', name='Go', **BOX).vba('Click', 'Beep')
        components = sheet.ole.project.components
        self.assertIn('Private Sub Go_Click()',
                      components['Sheet1'].CodeModule().text)
        self.assertNotIn('WineOLE', components,
                         "a worksheet ActiveX handler belongs in the sheet's "
                         'own module and nowhere else')


if __name__ == '__main__':
    unittest.main()
