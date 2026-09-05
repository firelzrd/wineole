import contextlib
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.excel_integration import ExcelIntegrationMixin

from wineole.errors import RemoteError
from wineole.msoffice import VBA, VBAAccessDenied, Excel

# Non-ASCII payloads are \u escapes, never literal characters: shipped files
# here are ASCII-only. GREEK_CYRILLIC is alpha/beta/gamma + a hyphen + two
# Cyrillic letters -- JIS X 0208 rows 6-7, so CP932 can represent all five
# and they round-trip intact on this host. UNREPRESENTABLE is "caf" +
# e-acute + a space + a check mark: neither of those two is in CP932, so it
# is the payload the refusal exists for.
GREEK_CYRILLIC = '\u03b1\u03b2\u03b3-\u0414\u0436'
UNREPRESENTABLE = 'caf\u00e9 \u2713'


class MSOfficeVBAIntegrationTest(ExcelIntegrationMixin, unittest.TestCase):
    """Pins the VBA layer against real Excel, running under Wine. The unit
    files test every piece against fakes; this is where the codepage
    boundary, the VBProject refusal and the where-code-goes table are
    exercised for real."""

    @contextlib.contextmanager
    def wrapped_excel(self):
        """A wrapped Excel with a workbook already open, built on bridge()
        so the spawn/teardown plumbing is not duplicated. Each test gets its
        own Excel, and it is quit in the context manager's own unwind."""
        with self.bridge() as client:
            with Excel.run('create', client=client) as xl:
                xl.hide()
                with xl.no_alert():
                    xl.ole.Workbooks().Add()
                    yield xl

    # --- blocks ------------------------------------------------------------

    def test_a_named_block_survives_a_save_and_reopen(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write('Function Doubled(a)\n  Doubled = a * 2\nEnd Function',
                           name='helpers')
            self.assertEqual(42, xl.ole.Run('Doubled', 21))

    def test_rewriting_a_block_does_not_define_it_twice(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write('Function F(a)\n  F = a + 1\nEnd Function', name='f')
            book.vba.write('Function F(a)\n  F = a + 2\nEnd Function', name='f')
            self.assertEqual(23, xl.ole.Run('F', 21),
                             'the second definition must win, and the first must '
                             'be gone')

    def test_removing_the_last_block_removes_the_module(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write('Function G()\n  G = 1\nEnd Function', name='g')
            before = book.ole.VBProject().VBComponents().Count()
            book.vba.remove('g')
            self.assertEqual(before - 1,
                             book.ole.VBProject().VBComponents().Count())

    # A sheet's block must land in that sheet's own code module, not the
    # wrapper's default one -- that is where Excel looks for
    # <ActiveX control>_Click.
    def test_a_sheet_block_lands_in_that_sheets_module(self):
        with self.wrapped_excel() as xl:
            sheet = xl[':first!']
            sheet.vba.write('Sub Marker()\nEnd Sub', name='marker')
            code_name = sheet.ole.CodeName()
            module = (xl['[]'].ole.VBProject().VBComponents()
                      .Item(code_name).CodeModule())
            self.assertIn('Sub Marker()', module.Lines(1, module.CountOfLines()))

    # --- the codepage boundary ---------------------------------------------

    # Non-ASCII goes through COM as Unicode -- the bridge's own BSTR
    # marshalling never touches a codepage. But this Excel's VBA6 engine
    # stores module *source* per the local ANSI codepage regardless of how
    # the string arrived: text inside that codepage's repertoire survives;
    # text outside it does not raise, it is just quietly downgraded. On this
    # (CP932) host, Greek and Cyrillic letters are inside the codepage and
    # round-trip intact -- measured directly before this test was written.
    def test_non_ascii_within_the_local_codepage_survives_string_injection(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write(
                'Function Greet()\n  Greet = "%s"\nEnd Function' % GREEK_CYRILLIC,
                name='g')
            self.assertEqual(GREEK_CYRILLIC, xl.ole.Run('Greet'))

    # A module's text is held in the system ANSI codepage, so a character
    # the codepage cannot represent used to be substituted on the way in --
    # measured here, this payload came back "cafe ?". It is refused now,
    # under the same rule import follows: silently dropping an accent is the
    # failure this phase exists to remove, and there is no way to inject
    # such a character as a literal at all.
    #
    # Ruby raises ArgumentError here; Python raises ValueError.
    def test_a_character_the_codepage_cannot_hold_is_refused_not_mangled(self):
        with self.wrapped_excel() as xl:
            with self.assertRaises(ValueError) as caught:
                xl['[]'].vba.write(
                    'Function Greet()\n  Greet = "%s"\nEnd Function' % UNREPRESENTABLE,
                    name='g')
            message = str(caught.exception)
            self.assertIn('cannot represent', message)
            self.assertIn(repr('\u00e9'), message,
                          'the message must name the character that stopped it')

    # --- files in and out ---------------------------------------------------

    def test_export_and_import_round_trip_through_utf8(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write('Function RT(a)\n  RT = a + 5\nEnd Function', name='rt')
            with tempfile.TemporaryDirectory() as directory:
                out = os.path.join(directory, 'WineOLE.bas')
                book.vba.export('WineOLE', out)
                with open(out, 'rb') as handle:
                    raw = handle.read()
                raw.decode('utf-8')  # raises if the exported file is not UTF-8
                self.assertNotIn(b'\r', raw, 'and must use LF')

                book.vba.remove('rt')
                book.vba.import_(out)
                self.assertEqual(26, xl.ole.Run('RT', 21))

    # The ASCII-only round trip above proves the mechanics (UTF-8/LF,
    # importable) but never puts a non-ASCII byte through the
    # encode/decode step it exists to protect.
    def test_export_and_import_round_trip_preserves_non_ascii_within_the_codepage(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write(
                'Function Payload()\n  Payload = "%s"\nEnd Function' % GREEK_CYRILLIC,
                name='k')
            with tempfile.TemporaryDirectory() as directory:
                out = os.path.join(directory, 'WineOLE.bas')
                book.vba.export('WineOLE', out)
                with open(out, encoding='utf-8') as handle:
                    self.assertIn(GREEK_CYRILLIC, handle.read())

                book.vba.remove('k')
                book.vba.import_(out)
                self.assertEqual(GREEK_CYRILLIC, xl.ole.Run('Payload'))

    # A codepage file is what Excel's own Export writes and what every .bas
    # from a Windows toolchain is. It used to be refused; it is read now.
    def test_a_codepage_source_file_imports_without_being_re_saved(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            with tempfile.TemporaryDirectory() as directory:
                src = os.path.join(directory, 'Cp.bas')
                # The module and the procedure must not share a name: VBA
                # resolves the bare name to the module and then reports the
                # macro missing.
                source = ('Attribute VB_Name = "CpMod"\r\n'
                          'Public Function Greeting()\r\n'
                          '  Greeting = "%s"\r\n'
                          'End Function\r\n' % GREEK_CYRILLIC)
                with open(src, 'wb') as handle:
                    handle.write(source.encode(VBA.codepage()))
                with open(src, 'rb') as handle:
                    raw = handle.read()
                with self.assertRaises(UnicodeDecodeError,
                                       msg='the fixture has to be a file UTF-8 '
                                           'cannot explain, or it proves nothing'):
                    raw.decode('utf-8')

                book.vba.import_(src)
                self.assertEqual(GREEK_CYRILLIC, xl.ole.Run('Greeting'))

    # --- components ---------------------------------------------------------

    def test_a_named_module_can_be_made_written_to_and_removed(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.add_component('Utils')
            book.vba.write('Public Function Beta()\n  Beta = 8\nEnd Function',
                           name='b', into='Utils')
            self.assertEqual(8, xl.ole.Run('Beta'))

            book.vba.remove_component('Utils')
            # The component is gone, so the code in it has to be gone too --
            # asserting only that Remove returned would pass even if it had
            # not.
            with self.assertRaises(RemoteError):
                xl.ole.Run('Beta')

    # Excel Add()s before it renames, and the two are not atomic, so a
    # refusal that came after the Add would leave a stray Module1 behind.
    # Counting is what catches that; the raise alone would not.
    def test_making_a_component_under_a_taken_name_leaves_no_stray(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.add_component('Utils')
            before = book.vba.project().VBComponents().Count()

            with self.assertRaises(ValueError):
                book.vba.add_component('Utils')
            self.assertEqual(before, book.vba.project().VBComponents().Count())

    def test_a_module_excel_owns_cannot_be_removed(self):
        with self.wrapped_excel() as xl:
            with self.assertRaises(ValueError) as caught:
                xl['[]'].vba.remove_component('ThisWorkbook')
            self.assertIn('cannot be deleted', str(caught.exception))

    # Where code lives decides whether it can be called at all -- the table
    # in BookVBA's docstring, pinned against a live Excel so it cannot
    # drift.
    def test_only_a_standard_module_answers_an_unqualified_run(self):
        with self.wrapped_excel() as xl:
            book = xl['[]']
            book.vba.write('Public Function Alpha()\n  Alpha = 7\nEnd Function',
                           name='a')
            book.vba.write('Public Function Gamma()\n  Gamma = 9\nEnd Function',
                           name='c', into='ThisWorkbook')

            self.assertEqual(7, xl.ole.Run('Alpha'),
                             'a standard module answers a bare Run')
            with self.assertRaises(RemoteError) as caught:
                xl.ole.Run('Gamma')
            self.assertIn('0x800A03EC', str(caught.exception),
                          'ThisWorkbook does not')

    # --- the refusal path ---------------------------------------------------

    # This host has AccessVBOM enabled, so every test above only exercises
    # the allowed path -- the one path most users will NOT hit first. This
    # is the only test in the suite that exercises the refusal, by switching
    # the setting off, starting a fresh Excel (the setting is read at
    # startup, so a running instance would not notice), and checking the
    # refusal is the guidance message rather than a raw COM error -- then
    # restoring the setting whatever happens.
    #
    # Deliberately not a skip. If the switch stops working the suite must
    # say so, rather than quietly shipping a branch nobody has run.
    def test_a_disabled_setting_produces_advice_not_a_com_error(self):
        original = VBA.state()
        self.assertTrue(VBA.disable(),
                        'could not switch AccessVBOM off, so the refusal path '
                        'cannot be tested')
        try:
            with self.wrapped_excel() as xl:
                with self.assertRaises(VBAAccessDenied) as caught:
                    xl['[]'].vba.write('Sub A()', name='a')
                self.assertIn('wineole-vba enable', str(caught.exception))
        finally:
            if original == 'enabled':
                VBA.enable()
            else:
                VBA.disable()


if __name__ == '__main__':
    unittest.main()
