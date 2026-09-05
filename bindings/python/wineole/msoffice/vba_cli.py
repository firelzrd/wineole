"""The `wineole-vba` command: read and set the AccessVBOM registry switch.

Imports `wineole.msoffice.vba` and nothing else -- not `wineole`, not
`wineole.msoffice`. The registry read/write is a local `wine reg`
subprocess, never the JSON-Lines bridge, so this command needs no bridge
running and must never open a connection.
"""

import sys

from .vba import VBA

USAGE = """usage: wineole-vba [status|enable|disable]

  status    show whether programmatic access to VBA projects is on
  enable    turn it on
  disable   turn it off

Excel reads this setting when it starts, so an Excel that is already
running keeps whatever it had. Restart it for a change to take effect.

This is a macro security setting. Turning it on lets any Office
automation on this machine reach VBA projects, not just this library.
"""


def report(state):
    if state == 'enabled':
        print('VBA project access: enabled')
    elif state == 'disabled':
        print('VBA project access: disabled')
    elif state == 'unset':
        print('VBA project access: disabled (the registry value is not set)')


def main(argv=None):
    """Returns the exit code rather than calling sys.exit, so a test can
    call it directly and read the number. The module's __main__ block and
    the console-script entry point are what turn it into an exit status.
    """
    if argv is None:
        argv = sys.argv[1:]

    # No argument and 'status' are the same command. Extra arguments after a
    # valid subcommand are ignored, exactly as the Ruby script ignores
    # ARGV[1..].
    command = argv[0] if argv else 'status'

    if command == 'status':
        report(VBA.state())
        return 0

    if command in ('enable', 'disable'):
        ok = VBA.enable() if command == 'enable' else VBA.disable()
        if not ok:
            print('wineole-vba: could not write the registry (is wine on PATH?)',
                  file=sys.stderr)
            return 1
        report(VBA.state())
        print('Restart Excel for this to take effect -- it reads the setting '
              'at startup.')
        return 0

    # end='' because USAGE already ends with its own newline.
    print(USAGE, end='', file=sys.stderr)
    return 2


if __name__ == '__main__':
    sys.exit(main())
