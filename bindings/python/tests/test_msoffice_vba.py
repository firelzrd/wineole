import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from wineole.msoffice.vba import VBA, VBAError


class MSOfficeVBATest(unittest.TestCase):
    """Every test here replaces VBA.run_reg, the one method that shells out,
    so no test touches the real registry.

    The swap is `VBA.run_reg = staticmethod(fake)`, not a plain function:
    the production code calls `cls.run_reg(args)`, and a staticmethod is the
    one descriptor that makes that call arrive as `fake(args)` with nothing
    prepended, whatever the attribute is reached through. The original is
    taken from `VBA.__dict__` (the undecorated classmethod object) so
    putting it back restores exactly what was there.
    """

    def setUp(self):
        # codepage memoizes, so without this a test that simulates a broken
        # registry would read a value an earlier test had already cached and
        # pass or fail on test order rather than on what it is testing.
        VBA.forget_codepage()
        self.original_run_reg = VBA.__dict__['run_reg']

    def tearDown(self):
        VBA.run_reg = self.original_run_reg
        VBA.forget_codepage()

    def stub_reg(self, result):
        """Replace run_reg with something answering `result`. Returns the
        list the argv of every call is appended to."""
        calls = []

        def fake(args):
            calls.append(list(args))
            return result

        VBA.run_reg = staticmethod(fake)
        return calls

    def test_state_reads_the_dword(self):
        self.stub_reg(('\nHKEY_CURRENT_USER\\...\n    AccessVBOM    REG_DWORD    0x1\n', True))
        self.assertEqual('enabled', VBA.state())
        self.assertIs(True, VBA.enabled())

    def test_state_when_disabled(self):
        self.stub_reg(('    AccessVBOM    REG_DWORD    0x0\n', True))
        self.assertEqual('disabled', VBA.state())
        self.assertIs(False, VBA.enabled())

    # A missing value and a missing key both exit non-zero, and wine writes
    # the explanation to stdout in its own language -- so only the exit
    # status can be trusted.
    def test_state_when_the_value_is_absent(self):
        self.stub_reg(('reg: <a localized not-found message>', False))
        self.assertEqual('unset', VBA.state())
        self.assertIs(False, VBA.enabled())

    # Measured: wine indents the line and leaves a CR on the end, so the raw
    # field is neither "932" nor at the index you would expect. Stripping the
    # whole line before splitting is what fixes both -- drop that and this
    # test fails, which is how it was confirmed to be the load-bearing part.
    def test_the_codepage_is_read_from_the_registry(self):
        self.stub_reg(('    ACP    REG_SZ    932\r\n', True))
        self.assertEqual('cp932', VBA.codepage())

    # `read` picks its line by matching the NAME field exactly, not by
    # substring -- so a value whose *data* happens to contain another value's
    # name must not be mistaken for that value's line. "NameAccessVBOM"
    # contains the value name; mistaking this line for the real one would
    # hand "NameAccessVBOM" to int(..., 16), which raises -- while the real
    # line says enabled.
    def test_state_is_not_fooled_by_the_name_appearing_inside_other_data(self):
        self.stub_reg(('    SomethingElse    REG_SZ    NameAccessVBOM\r\n'
                       '    AccessVBOM    REG_DWORD    0x1\r\n', True))
        self.assertEqual('enabled', VBA.state())

    # More than one line naming the value means the single-line shape this
    # parser assumes has stopped holding -- a bug to surface, not a state to
    # quietly resolve by taking the first match.
    def test_state_raises_when_more_than_one_line_names_the_value(self):
        self.stub_reg(('    AccessVBOM    REG_DWORD    0x1\r\n'
                       '    AccessVBOM    REG_DWORD    0x0\r\n', True))
        with self.assertRaises(VBAError):
            VBA.state()

    # The command can succeed (exit 0) yet produce output this parser cannot
    # make sense of -- that is a parser bug or a wine output change, and must
    # not be reported as 'unset' (which means "never configured").
    def test_state_raises_rather_than_reporting_unset_when_output_does_not_parse(self):
        self.stub_reg(('some unrelated line that never names the value\r\n', True))
        with self.assertRaises(VBAError):
            VBA.state()

    def test_state_is_unset_only_when_the_command_says_the_value_is_not_there(self):
        self.stub_reg(('reg: <a localized not-found message>', False))
        self.assertEqual('unset', VBA.state())

    def test_an_unreadable_codepage_raises_rather_than_guessing(self):
        self.stub_reg(('reg: <not found>', False))
        with self.assertRaises(VBAError) as caught:
            VBA.codepage()
        self.assertIn('ACP', str(caught.exception))

    def test_an_unknown_codepage_raises(self):
        self.stub_reg(('    ACP    REG_SZ    99999\r\n', True))
        with self.assertRaises(VBAError):
            VBA.codepage()

    def test_enable_and_disable_pass_the_right_arguments(self):
        calls = self.stub_reg(('', True))
        self.assertIs(True, VBA.enable())
        self.assertEqual(1, len(calls))
        args = calls[0]
        self.assertEqual('add', args[0])
        self.assertIn('AccessVBOM', args)
        self.assertIn('REG_DWORD', args)
        self.assertEqual('1', args[args.index('/d') + 1])

        calls = self.stub_reg(('', True))
        VBA.disable()
        args = calls[0]
        self.assertEqual('0', args[args.index('/d') + 1])

    # wine reg add prints a localized success message to stdout. Only the
    # exit status says whether it worked.
    def test_enable_reports_failure_by_exit_status_not_by_output(self):
        self.stub_reg(('<some localized text>', True))
        self.assertIs(True, VBA.enable())
        self.stub_reg(('<some localized text>', False))
        self.assertIs(False, VBA.enable())

    # codepage costs a `wine reg` subprocess -- 328 ms measured on this host
    # -- and an import used to pay it twice. The machine's ANSI codepage
    # cannot change while a process runs, so reading it once is not a
    # staleness risk.
    def test_the_codepage_is_read_from_the_registry_only_once(self):
        calls = self.stub_reg(('    ACP    REG_SZ    932\r\n', True))
        first = VBA.codepage()
        self.assertEqual(first, VBA.codepage())
        self.assertEqual(first, VBA.codepage())
        self.assertEqual(1, len(calls), 'three calls, one registry read')


if __name__ == '__main__':
    unittest.main()
