import contextlib
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import (FakeClient, FakeComWorkbook,
                              FakeComWorksheetWithProject,
                              FakeProjectWhoseLookupFails, FakeVBProject)

from wineole.errors import RemoteError
from wineole.msoffice.book import Book
from wineole.msoffice.sheet import Sheet
from wineole.msoffice.vba import VBA, VBAAccessDenied, VBAError


def make_book(ole=None, loopback=True, convert_paths=True,
              vb_project=None, denied=False):
    if ole is None:
        ole = FakeComWorkbook(vb_project=vb_project, vb_project_denied=denied)
    return Book(ole, FakeClient(loopback=loopback), '11.0',
                convert_paths=convert_paths)


def make_sheet(vb_project=None, denied=False):
    return Sheet(FakeComWorksheetWithProject('Sheet1', vb_project=vb_project,
                                             denied=denied), '11.0')


@contextlib.contextmanager
def with_codepage(name):
    """Pin the codepage without a registry. Swaps the classmethod for a
    staticmethod so `VBA.codepage()` and `cls.codepage()` both land on the
    stub, and puts the original descriptor back afterwards."""
    original = VBA.__dict__['codepage']
    VBA.codepage = staticmethod(lambda: name)
    try:
        yield
    finally:
        VBA.codepage = original


@contextlib.contextmanager
def with_reg(result):
    """VBA.denied consults the real registry to pick its wording, so a test
    of the denial path would otherwise pass or fail on whatever AccessVBOM
    happens to be on the host running the suite."""
    original = VBA.__dict__['run_reg']
    VBA.run_reg = staticmethod(lambda args: result)
    VBA.forget_codepage()
    try:
        yield
    finally:
        VBA.run_reg = original
        VBA.forget_codepage()


NOT_FOUND = ('reg: <a localized not-found message>', False)
ENABLED = ('    AccessVBOM    REG_DWORD    0x1\r\n', True)


class MSOfficeBookVBATest(unittest.TestCase):

    # --- blocks in the wrapper's own module -------------------------------

    def test_vba_puts_a_named_block_in_the_wrappers_own_module(self):
        project = FakeVBProject()
        make_book(vb_project=project).vba.write('Sub Go()\nEnd Sub', name='go')
        module = project.components['WineOLE']
        self.assertIn("'<wineole:go>", module.CodeModule().text)
        self.assertIn('Sub Go()', module.CodeModule().text)

    def test_removing_the_last_block_removes_the_module(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        book.vba.write('Sub Go()', name='go')
        book.vba.remove('go')
        self.assertNotIn('WineOLE', project.components,
                         'an empty module is litter -- it should go when its '
                         'last block does')

    def test_removing_one_of_two_blocks_keeps_the_module(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        book.vba.write('Sub A()', name='a')
        book.vba.write('Sub B()', name='b')
        book.vba.remove('a')
        self.assertIn('WineOLE', project.components)
        self.assertIn('Sub B()', project.components['WineOLE'].CodeModule().text)

    # remove asks VBABlock.remove to find and delete the block (one Lines
    # call) and must not fetch the body again afterwards just to ask whether
    # the module is now blank. Leftover content matters here: removing the
    # *only* block empties the module down to CountOfLines == 0, and the
    # body short-circuits on that without a Lines call either way -- which
    # would let this test pass by coincidence even with a double fetch.
    def test_remove_vba_fetches_the_body_once(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        book.vba.write('Sub A()', name='a')
        book.vba.write('Sub B()', name='b')
        module = project.components['WineOLE'].CodeModule()
        reads_before = module.reads

        book.vba.remove('a')

        self.assertEqual(reads_before + 1, module.reads,
                         'one Lines call, not one per operation')

    # --- a denied project gives advice, not a raw RemoteError -------------

    def test_a_denied_project_says_what_to_do(self):
        book = make_book(denied=True)
        with with_reg(NOT_FOUND):
            with self.assertRaises(VBAAccessDenied) as caught:
                book.vba.write('Sub A()', name='a')
        message = str(caught.exception)
        self.assertIn('wineole-vba enable', message)
        self.assertIn('restart Excel', message)

    # The second denial row: the registry is already enabled and access is
    # still refused, because a running Excel caches the setting from
    # startup. That case cannot be told apart from the first by HRESULT or
    # message text, so the message must consult VBA.state itself.
    def test_a_denied_project_with_the_registry_already_enabled_says_restart_excel(self):
        book = make_book(denied=True)
        with with_reg(ENABLED):
            with self.assertRaises(VBAAccessDenied) as caught:
                book.vba.write('Sub A()', name='a')
        message = str(caught.exception)
        self.assertIn('restart Excel', message)
        self.assertNotIn('wineole-vba enable', message,
                         'the registry is already enabled -- telling the reader '
                         'to enable it again is wrong')

    # --- into= targets an existing component, not the wrapper's own -------

    def test_into_targets_an_existing_component(self):
        project = FakeVBProject()
        project.add_existing('AppForm')
        book = make_book(vb_project=project)
        book.vba.write('Private Sub Go_Click()\nEnd Sub', name='go', into='AppForm')
        self.assertIn("'<wineole:go>",
                      project.components['AppForm'].CodeModule().text)
        self.assertNotIn('WineOLE', project.components,
                         'into= must not create the default module')

    # We clean up what we made, not what we were pointed at. A UserForm can
    # be deleted, unlike ThisWorkbook -- so the rule has to be stated, not
    # left to what COM happens to allow.
    #
    # A same-named block goes in the wrapper's own module *first*: without
    # it, a broken from_ branch falls through to the default path, finds no
    # "WineOLE" module at all, and does nothing whatsoever -- AppForm would
    # then survive by accident rather than by a working guard.
    def test_a_named_component_is_never_removed_even_when_it_empties(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        book.vba.write('Sub Go()', name='go')
        project.add_existing('AppForm')
        book.vba.write('Sub Go()', name='go', into='AppForm')

        book.vba.remove('go', from_='AppForm')

        self.assertIn('AppForm', project.components)
        self.assertNotIn('Sub Go()', project.components['AppForm'].CodeModule().text,
                         'the block must actually be removed from the named '
                         'component the call was for')
        self.assertIn('WineOLE', project.components,
                      'a call naming from_ must never fall through and touch '
                      "the wrapper's own module")
        self.assertIn("'<wineole:go>",
                      project.components['WineOLE'].CodeModule().text)

    def test_into_a_component_that_is_not_there_says_so(self):
        book = make_book()
        with self.assertRaises(ValueError) as caught:
            book.vba.write('Sub A()', name='a', into='Missing')
        self.assertIn('Missing', str(caught.exception))

    def test_remove_vba_from_a_component_that_is_not_there_says_so(self):
        book = make_book()
        with self.assertRaises(ValueError) as caught:
            book.vba.remove('a', from_='Missing')
        self.assertIn('Missing', str(caught.exception))

    # --- import_ / export --------------------------------------------------

    def test_export_converts_the_codepage_and_the_line_endings(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        project.add_existing('Mod')
        # What Excel writes: the ANSI codepage, CRLF.
        project.components['Mod'].export_bytes = (
            b'Attribute VB_Name = "Mod"\r\n\' caf\xe9\r\n')

        with tempfile.TemporaryDirectory() as directory:
            out = os.path.join(directory, 'Mod.bas')
            with with_codepage('cp1252'):
                book.vba.export('Mod', out)
            with open(out, 'rb') as handle:
                raw = handle.read()
            self.assertEqual(
                ('Attribute VB_Name = "Mod"\n\' caf\u00e9\n').encode('utf-8'),
                raw)
            self.assertNotIn(b'\r', raw,
                             'a file written to a Linux path should not carry CRLF')

    def test_import_converts_utf8_into_the_codepage(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            with open(src, 'wb') as handle:
                handle.write(
                    ('Attribute VB_Name = "Mod"\n\' caf\u00e9\n').encode('utf-8'))
            with with_codepage('cp1252'):
                book.vba.import_(src)
            self.assertIn(b"' caf\xe9", project.imported_bytes,
                          'Excel must be handed the codepage bytes, not UTF-8')

    # Silently substituting would manufacture the very failure this wrapper
    # exists to avoid: it succeeded, and the result is wrong.
    def test_a_character_the_codepage_cannot_hold_raises(self):
        book = make_book()
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            with open(src, 'wb') as handle:
                handle.write("' \U0001F600\n".encode('utf-8'))
            with with_codepage('cp1252'):
                with self.assertRaises(ValueError) as caught:
                    book.vba.import_(src)
            message = str(caught.exception)
            self.assertIn(repr('\U0001F600'), message,
                          'the character that stopped it')
            self.assertIn('cp1252', message)

    # The file is handed to Excel by path. On a remote bridge that path
    # means a different machine's filesystem, so there is nothing sensible
    # to do -- and it is the environment that is wrong, not the argument,
    # so this is RuntimeError rather than ValueError.
    def test_import_and_export_refuse_a_remote_bridge(self):
        book = make_book(loopback=False)
        with self.assertRaises(RuntimeError):
            book.vba.import_('/tmp/x.bas')
        with self.assertRaises(RuntimeError):
            book.vba.export('Mod', '/tmp/x.bas')

    # The write direction: a file that DECODES cleanly can still hold a
    # character the codepage cannot store, and a bare UnicodeEncodeError
    # names neither the file nor the way out -- while the string path gives
    # a full explanation for the very same condition. One message serves
    # both.
    def test_a_character_the_codepage_cannot_store_is_refused_by_name_on_the_file_path(self):
        book = make_book()
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            with open(src, 'wb') as handle:
                # "caf" + e-acute + space + check mark: CP932 holds
                # neither of the last two.
                handle.write("' caf\u00e9 \u2713\n".encode('utf-8'))
            with with_codepage('cp932'):
                with self.assertRaises(ValueError) as caught:
                    book.vba.import_(src)
            message = str(caught.exception)
            self.assertIn('cp932', message)
            self.assertIn(src, message, 'the message must name the file')
            self.assertIn('ChrW', message, 'and the way out')

    # Both paths are bound by the same codepage, so a caller must not be
    # able to tell from the message which one they hit -- only which
    # character stopped it. Two texts, one explanation.
    def test_the_string_path_and_the_file_path_explain_the_same_thing(self):
        book = make_book()
        with with_codepage('cp932'):
            with self.assertRaises(ValueError) as caught:
                book.vba.write('x = "caf\u00e9"', name='a')
            from_string = str(caught.exception)

            with tempfile.TemporaryDirectory() as directory:
                src = os.path.join(directory, 'Mod.bas')
                with open(src, 'wb') as handle:
                    handle.write('x = "caf\u00e9"\n'.encode('utf-8'))
                with self.assertRaises(ValueError) as caught:
                    book.vba.import_(src)
                from_file = str(caught.exception)

        marker = 'which the system codepage'
        self.assertEqual(from_string[from_string.index(marker):],
                         from_file[from_file.index(marker):])

    # --- add_component / remove_component ---------------------------------

    def test_add_component_creates_one_of_the_named_kind(self):
        book = make_book(vb_project=FakeVBProject())
        component = book.vba.add_component('Utils')
        self.assertEqual('Utils', component.Name())
        self.assertEqual(1, component.Type(), 'standard module by default')
        self.assertEqual(3, book.vba.add_component('Dialog', kind='form').Type())
        self.assertEqual(2, book.vba.add_component('Thing', kind='class').Type())

    # The name check has to happen BEFORE Add, not after. Excel's Add and
    # the rename that follows it are not atomic -- adding under a taken name
    # succeeds and only the rename fails, leaving a stray Module1 behind. So
    # the assertion that matters is not just "it raised" but "it added
    # nothing".
    def test_add_component_refuses_a_taken_name_without_adding_anything(self):
        project = FakeVBProject()
        project.add_existing('Utils')
        book = make_book(vb_project=project)

        with self.assertRaises(ValueError) as caught:
            book.vba.add_component('Utils')
        self.assertIn("already has a VBA component named 'Utils'",
                      str(caught.exception))
        self.assertEqual(1, len(project.component_list),
                         'nothing may be added on the refusal path')

    def test_add_component_refuses_an_unknown_kind_without_adding_anything(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)

        with self.assertRaises(ValueError) as caught:
            book.vba.add_component('Utils', kind='worksheet')
        self.assertIn("unknown component kind 'worksheet'", str(caught.exception))
        self.assertEqual([], project.component_list)

    # The pre-check cannot cover a name Excel itself rejects (too long, an
    # illegal character), so the Add can still succeed with the rename
    # failing after it. The stray has to be taken back out.
    def test_a_rename_that_fails_after_add_takes_the_stray_back_out(self):
        project = FakeVBProject()
        project.next_rename_fails = True
        book = make_book(vb_project=project)

        with self.assertRaises(ValueError) as caught:
            book.vba.add_component('this name is refused')
        self.assertIn('Excel refused', str(caught.exception))
        self.assertEqual([], project.component_list,
                         'the half-made component must not survive the failure')

    def test_remove_component_deletes_it(self):
        project = FakeVBProject()
        project.add_existing('Utils')
        book = make_book(vb_project=project)

        book.vba.remove_component('Utils')
        self.assertEqual([], project.component_list)

    # A worksheet's module and ThisWorkbook are Excel's, not ours. COM
    # refuses to remove them, and a raw refusal says nothing useful, so this
    # is caught before the call and answered with what to do instead.
    def test_remove_component_refuses_a_module_excel_owns(self):
        project = FakeVBProject()
        project.add_existing('ThisWorkbook', component_type=100)
        book = make_book(vb_project=project)

        with self.assertRaises(ValueError) as caught:
            book.vba.remove_component('ThisWorkbook')
        message = str(caught.exception)
        self.assertIn('cannot be deleted', message)
        self.assertIn('remove(name, from_=)', message,
                      'and must point at what does work')
        self.assertEqual(1, len(project.component_list),
                         'and must not have removed it')

    def test_remove_component_on_a_name_that_is_not_there_says_so(self):
        book = make_book(vb_project=FakeVBProject())
        with self.assertRaises(ValueError) as caught:
            book.vba.remove_component('Nope')
        self.assertIn("no VBA component named 'Nope'", str(caught.exception))

    # --- import decides the encoding on evidence, never on a guess ---------

    # The behaviour this replaced: a file already in the codepage used to be
    # REFUSED, on the reasoning that it was probably a mistake. It is not --
    # Excel's own Export writes the codepage, so every .bas from a Windows
    # toolchain arrives this way. Bytes that are not valid UTF-8 PROVE the
    # file is not UTF-8, and that is evidence, not a guess.
    def test_a_codepage_file_is_read_as_the_codepage_not_refused(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            # 0x92 is CP1252's curly apostrophe and is not valid UTF-8 in
            # any position, so UTF-8 is ruled out on evidence.
            with open(src, 'wb') as handle:
                handle.write(b"' caf\x92\n")
            with with_codepage('cp1252'):
                book.vba.import_(src)
            # Handed to Excel unchanged: it was already in the codepage
            # Excel wants, so the round trip through Unicode has to be
            # lossless.
            self.assertEqual(b"' caf\x92\n", project.imported_bytes)

    # A BOM is the one conclusive signal, so it wins over everything else --
    # including bytes that would otherwise be read as the codepage.
    def test_a_bom_decides_the_encoding_and_is_not_passed_through(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            with open(src, 'wb') as handle:
                handle.write(b"\xef\xbb\xbf' caf\xc3\xa9\n")
            with with_codepage('cp1252'):
                book.vba.import_(src)
            self.assertEqual(b"' caf\xe9\n", project.imported_bytes,
                             'decoded as UTF-8 and re-encoded to cp1252')
            self.assertFalse(project.imported_bytes.startswith(b'\xef\xbb\xbf'),
                             'the BOM must not reach the module text')

    # Both encodings wrong at once. UTF-8 is ruled out by the bytes, and the
    # codepage cannot read them either, so there is genuinely nothing to
    # infer -- and the error has to say that rather than surfacing a bare
    # UnicodeDecodeError that names neither the file nor why that encoding
    # was the one attempted.
    def test_bytes_valid_in_neither_encoding_say_so_instead_of_raising_a_bare_encoding_error(self):
        book = make_book()
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            # 0x92 leads a CP932 double-byte character; a bare \n cannot
            # follow it.
            with open(src, 'wb') as handle:
                handle.write(b"' caf\x92\n")
            with with_codepage('cp932'):
                with self.assertRaises(ValueError) as caught:
                    book.vba.import_(src)
            message = str(caught.exception)
            self.assertIn(src, message, 'the message must name the file')
            self.assertIn('cp932', message, 'and the encoding it tried')
            self.assertIn('not valid UTF-8', message,
                          'and why that encoding was the one tried')

    # An explicit encoding= skips detection entirely -- and when it is
    # wrong, the message must not blame a detection that never ran.
    def test_an_explicit_encoding_skips_detection_and_owns_the_failure(self):
        book = make_book()
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            with open(src, 'wb') as handle:
                handle.write(b"' caf\x92\n")
            with self.assertRaises(ValueError) as caught:
                book.vba.import_(src, encoding='cp932')
            message = str(caught.exception)
            self.assertIn('you passed encoding', message)
            self.assertNotIn('no BOM', message,
                             'detection did not run, so it must not be cited')

    def test_a_codepage_byte_pair_that_happens_to_be_valid_utf8_is_not_caught(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        with tempfile.TemporaryDirectory() as directory:
            src = os.path.join(directory, 'Mod.bas')
            # C2 A0 is CP1252's non-breaking space glued to itself becoming
            # U+00A0 when read as UTF-8 -- the file is genuinely valid UTF-8
            # by the only test available, so import proceeds.
            with open(src, 'wb') as handle:
                handle.write(b"' \xc2\xa0\n")
            with with_codepage('cp1252'):
                book.vba.import_(src)
            # Round-tripped as a single CP1252 byte (0xA0), not the two
            # original bytes -- the silent content change that cannot be
            # detected, left in place and pinned so it is not mistaken for a
            # regression.
            self.assertEqual(b"' \xa0\n", project.imported_bytes)

    # --- export's raise-don't-substitute side ------------------------------

    # Excel's exported bytes can contain something undefined in its own
    # reported codepage (0x81 has no character in CP1252), and decoding must
    # raise rather than substitute. Ruby surfaces the bare
    # Encoding::UndefinedConversionError here; Python wraps the
    # UnicodeDecodeError in VBAError so the component and the codepage are
    # named.
    def test_export_raises_when_the_reported_codepage_cannot_decode_the_bytes(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        project.add_existing('Mod')
        project.components['Mod'].export_bytes = (
            b'Attribute VB_Name = "Mod"\r\n\' \x81\r\n')

        with tempfile.TemporaryDirectory() as directory:
            out = os.path.join(directory, 'Mod.bas')
            with with_codepage('cp1252'):
                with self.assertRaises(VBAError) as caught:
                    book.vba.export('Mod', out)
            message = str(caught.exception)
            self.assertIn('Mod', message)
            self.assertIn('cp1252', message)

    # --- a path ending in ".." gets a clear error, not a bare OSError ------

    def test_import_a_path_ending_in_dotdot_says_so_clearly(self):
        book = make_book()
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError) as caught:
                book.vba.import_(os.path.join(directory, '..'))
            self.assertIn('..', str(caught.exception))

    def test_export_a_path_ending_in_dotdot_says_so_clearly(self):
        project = FakeVBProject()
        book = make_book(vb_project=project)
        project.add_existing('Mod')
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(ValueError) as caught:
                book.vba.export('Mod', os.path.join(directory, '..'))
            self.assertIn('..', str(caught.exception))


class MSOfficeSheetVBATest(unittest.TestCase):

    # The denial rescue used to wrap the whole lookup, so ANY COM failure
    # after the project opened came back as "turn on AccessVBOM" -- advice
    # for a condition the caller is not in. Only the VBProject fetch is the
    # denial.
    def test_a_failure_after_the_project_opens_is_not_reported_as_access_denied(self):
        sheet = make_sheet(vb_project=FakeProjectWhoseLookupFails())
        with self.assertRaises(RemoteError) as caught:
            sheet.vba.write('Sub A()', name='a')
        self.assertIn('0x800A0009', str(caught.exception))
        self.assertNotIsInstance(caught.exception, VBAAccessDenied,
                                 'a missing component is not a permissions problem')

    def test_sheet_vba_writes_into_this_sheets_own_module(self):
        project = FakeVBProject()
        project.add_existing('Sheet1')
        sheet = make_sheet(vb_project=project)
        sheet.vba.write('Private Sub Go_Click()\nEnd Sub', name='go')
        self.assertIn("'<wineole:go>",
                      project.components['Sheet1'].CodeModule().text)

    def test_sheet_remove_vba_leaves_the_module_in_place(self):
        project = FakeVBProject()
        project.add_existing('Sheet1')
        sheet = make_sheet(vb_project=project)
        sheet.vba.write('Sub Go()', name='go')
        sheet.vba.remove('go')
        self.assertIn('Sheet1', project.components,
                      "a sheet's module cannot be removed, and must not be attempted")
        self.assertNotIn('Sub Go()', project.components['Sheet1'].CodeModule().text)

    def test_sheet_vba_on_a_denied_workbook_says_what_to_do(self):
        sheet = make_sheet(denied=True)
        with with_reg(NOT_FOUND):
            with self.assertRaises(VBAAccessDenied) as caught:
                sheet.vba.write('Sub A()', name='a')
        message = str(caught.exception)
        self.assertIn('wineole-vba enable', message)
        self.assertIn('restart Excel', message)

    def test_sheet_remove_vba_on_a_denied_workbook_says_what_to_do(self):
        sheet = make_sheet(denied=True)
        with with_reg(NOT_FOUND):
            with self.assertRaises(VBAAccessDenied) as caught:
                sheet.vba.remove('a')
        message = str(caught.exception)
        self.assertIn('wineole-vba enable', message)
        self.assertIn('restart Excel', message)

    # The registry-already-enabled branch: access is still refused because a
    # running Excel caches the setting from startup, so the message must
    # steer toward restarting Excel instead of re-running `wineole-vba
    # enable`.
    def test_sheet_vba_on_a_denied_workbook_with_the_registry_already_enabled_says_restart_excel(self):
        sheet = make_sheet(denied=True)
        with with_reg(ENABLED):
            with self.assertRaises(VBAAccessDenied) as caught:
                sheet.vba.write('Sub A()', name='a')
        message = str(caught.exception)
        self.assertIn('restart Excel', message)
        self.assertNotIn('wineole-vba enable', message,
                         'the registry is already enabled -- telling the reader '
                         'to enable it again is wrong')


if __name__ == '__main__':
    unittest.main()
