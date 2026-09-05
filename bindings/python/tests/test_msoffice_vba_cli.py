import contextlib
import io
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from wineole.msoffice.vba import VBA
from wineole.msoffice.vba_cli import USAGE, main

ENABLED = ('    AccessVBOM    REG_DWORD    0x1\r\n', True)
DISABLED = ('    AccessVBOM    REG_DWORD    0x0\r\n', True)
NOT_FOUND = ('reg: <a localized not-found message>', False)


class MSOfficeVBACLITest(unittest.TestCase):
    """The CLI never spawns wine: run_reg is replaced here exactly as it is
    in test_msoffice_vba.py, and stdout/stderr are captured so the exact
    lines can be asserted rather than eyeballed."""

    def setUp(self):
        self.original_run_reg = VBA.__dict__['run_reg']
        VBA.forget_codepage()

    def tearDown(self):
        VBA.run_reg = self.original_run_reg
        VBA.forget_codepage()

    def stub_reg(self, result):
        calls = []

        def fake(args):
            calls.append(list(args))
            return result

        VBA.run_reg = staticmethod(fake)
        return calls

    def run_cli(self, argv):
        """Returns (exit_code, stdout, stderr)."""
        out = io.StringIO()
        err = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            code = main(argv)
        return code, out.getvalue(), err.getvalue()

    # No argument and 'status' are the same command, so the no-argument form
    # is what this asserts -- it is the one a person types.
    def test_status_reports_an_enabled_setting(self):
        self.stub_reg(ENABLED)
        code, out, err = self.run_cli([])
        self.assertEqual(0, code)
        self.assertEqual('VBA project access: enabled\n', out)
        self.assertEqual('', err)

    def test_status_reports_a_disabled_setting(self):
        self.stub_reg(DISABLED)
        code, out, err = self.run_cli(['status'])
        self.assertEqual(0, code)
        self.assertEqual('VBA project access: disabled\n', out)
        self.assertEqual('', err)

    # 'unset' is still disabled in effect, but saying so differently is what
    # tells a reader that nothing has ever been configured here.
    def test_status_reports_an_absent_value_as_disabled_and_says_so(self):
        self.stub_reg(NOT_FOUND)
        code, out, err = self.run_cli(['status'])
        self.assertEqual(0, code)
        self.assertEqual(
            'VBA project access: disabled (the registry value is not set)\n', out)
        self.assertEqual('', err)

    # The restart line is not boilerplate: Excel reads the setting once, at
    # startup, so a running Excel keeps whatever it had.
    def test_enable_prints_the_new_state_and_the_restart_reminder(self):
        calls = self.stub_reg(ENABLED)
        code, out, err = self.run_cli(['enable'])
        self.assertEqual(0, code)
        self.assertEqual(
            'VBA project access: enabled\n'
            'Restart Excel for this to take effect -- it reads the setting at '
            'startup.\n', out)
        self.assertEqual('', err)
        self.assertEqual('add', calls[0][0])
        self.assertEqual('1', calls[0][calls[0].index('/d') + 1])

    def test_enable_that_cannot_write_the_registry_exits_1_with_the_reason(self):
        self.stub_reg(('', False))
        code, out, err = self.run_cli(['enable'])
        self.assertEqual(1, code)
        self.assertEqual('', out)
        self.assertEqual(
            'wineole-vba: could not write the registry (is wine on PATH?)\n', err)

    def test_an_unknown_subcommand_prints_usage_to_stderr_and_exits_2(self):
        self.stub_reg(ENABLED)
        code, out, err = self.run_cli(['wat'])
        self.assertEqual(2, code)
        self.assertEqual('', out)
        self.assertEqual(USAGE, err)


if __name__ == '__main__':
    unittest.main()
