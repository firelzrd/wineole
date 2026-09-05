"""Programmatic access to a workbook's VBA project is off by default, and
turning it on means writing a macro security setting in the registry.

This module can read and write it -- but nothing in the wrapper calls the
writing half. That is for `wineole-vba`, which a human runs, and for tests,
which need to exercise both sides of the switch.

It also owns two things that sit "below" the VBA object model and are
needed by everything above it: the machine's ANSI codepage (memoized, since
reading it costs a subprocess) and the encoding rules a .bas file is read
by. It imports nothing from the rest of `msoffice`, so the CLI can load it
without the composition root and without a bridge.
"""

import codecs
import subprocess

from ..errors import WineOLEError


class VBAError(WineOLEError):
    pass


class VBAAccessDenied(VBAError):
    pass


class VBA:
    """A class used only through its classmethods; never instantiated.

    Ruby's `module VBA` with `self.` methods and one `@codepage` memo has no
    exact Python shape. A class with classmethods and a class attribute is
    the closest: the memo is real state with a name, and `forget_codepage`
    resets it the way the Ruby test seam does.
    """

    ACCESS_KEY = 'HKCU\\Software\\Microsoft\\Office\\11.0\\Excel\\Security'
    ACCESS_VALUE = 'AccessVBOM'
    CODEPAGE_KEY = 'HKLM\\System\\CurrentControlSet\\Control\\Nls\\CodePage'

    # BOMs, longest first -- FF FE 00 00 is UTF-32LE and also starts with
    # the UTF-16LE BOM, so a shorter match must never be tried first.
    BOMS = [
        (b'\x00\x00\xfe\xff', 'utf-32-be'),
        (b'\xff\xfe\x00\x00', 'utf-32-le'),
        (b'\xef\xbb\xbf', 'utf-8'),
        (b'\xfe\xff', 'utf-16-be'),
        (b'\xff\xfe', 'utf-16-le'),
    ]

    # So `VBA.Error` and `VBA.AccessDenied` read the way Ruby's nested
    # classes do, without making the module-level names harder to import.
    Error = VBAError
    AccessDenied = VBAAccessDenied

    _codepage = None

    @classmethod
    def state(cls):
        """'enabled', 'disabled', or 'unset' when the value is not there."""
        raw = cls.read(cls.ACCESS_KEY, cls.ACCESS_VALUE)
        if raw is None:
            return 'unset'
        return 'disabled' if int(raw, 16) == 0 else 'enabled'

    @classmethod
    def enabled(cls):
        return cls.state() == 'enabled'

    @classmethod
    def denied(cls):
        """Never returns.

        Touching VBProject when access is refused gives 0x800A03EC and a
        localized message -- the same HRESULT a rejected NumberFormat gives,
        and neither identifies the condition. The registry is what turns the
        refusal into advice.
        """
        if cls.state() == 'enabled':
            message = (
                'access to the VBA project was refused even though the registry '
                'has it enabled -- Excel reads that setting when it starts, so '
                'restart Excel if it was switched on while this instance was running')
        else:
            message = (
                'access to the VBA project is disabled. Run `wineole-vba enable`, '
                'then restart Excel -- it reads the setting at startup')
        raise VBAAccessDenied(message)

    @classmethod
    def enable(cls):
        return cls.write(cls.ACCESS_KEY, cls.ACCESS_VALUE, '1')

    @classmethod
    def disable(cls):
        return cls.write(cls.ACCESS_KEY, cls.ACCESS_VALUE, '0')

    @classmethod
    def codepage(cls):
        """The Windows ANSI codepage of this prefix, as a codec name Python
        knows (e.g. 'cp932'). Never hardcoded: VBA source files are written
        and read in whatever this is, and it is not CP932 everywhere.

        Memoized because it costs a `wine reg` subprocess -- measured at
        328 ms on this host -- and a machine's ANSI codepage does not change
        while a process runs. Without this an import pays it twice and a
        non-ASCII refusal pays it three times.
        """
        if cls._codepage is None:
            cls._codepage = cls.read_codepage()
        return cls._codepage

    @classmethod
    def forget_codepage(cls):
        """For tests, which swap the codepage to exercise both sides of it."""
        cls._codepage = None
        return None

    @classmethod
    def read_codepage(cls):
        raw = cls.read(cls.CODEPAGE_KEY, 'ACP')
        if raw is None:
            raise VBAError(
                f"could not read the ANSI codepage (ACP) from {cls.CODEPAGE_KEY}")

        name = 'cp' + raw
        try:
            codecs.lookup(name)
        except LookupError:
            raise VBAError(
                f"the registry reports ANSI codepage {raw!r}, which Python "
                'does not know') from None
        return name

    @classmethod
    def unrepresentable(cls, char, where):
        """Never returns.

        One explanation, used by both paths that hand text to Excel. They
        are bound by the same codepage and fail the same way, so the caller
        should not be able to tell from the message which one they hit --
        only which character stopped it.

        `where` says what the text was, because the way out differs: code
        given as a string can be rewritten with ChrW(), a file has to be
        edited.
        """
        raise ValueError(
            f"{where} contains {char!r}, which the system codepage "
            f"({cls.codepage()}) cannot represent. Excel stores a module's text "
            'in that codepage, so the character would be silently replaced '
            'rather than stored. Rewrite it with Chr()/ChrW() escapes, which '
            'are built at run time and are not bound by the codepage the '
            'source text is')

    @classmethod
    def detect_encoding(cls, path):
        """What encoding a VBA source file should be read as. Three rules,
        and every one of them decides on evidence rather than on a guess.

          1. A BOM is conclusive. Follow it.
          2. Bytes that are not valid UTF-8 PROVE the file is not UTF-8, so
             read it as the ANSI codepage -- which is what Excel's own
             Export writes, and what every .bas from a Windows toolchain is.
          3. Otherwise UTF-8.

        Rule 2 is the one that earns its place, and the direction matters:
        measured on this host, a CP932 file read as UTF-8 is invalid 95.07%
        of the time at ONE non-ASCII character and 99.99% by five, so real
        codepage files land here almost without exception. The reverse does
        not hold -- UTF-8 bytes read as CP932 come out VALID from two
        characters on, silently wrong. That asymmetry is why UTF-8 is the
        fallback in rule 3 and the codepage is never the default: guessing
        UTF-8 and being wrong is loud, guessing the codepage and being wrong
        is silent.

        What is left is a file in the codepage whose bytes happen to be
        valid UTF-8 -- undecidable, by construction, for anyone. Pass
        `encoding=` to skip all of this when you already know.
        """
        with open(path, 'rb') as handle:
            head = handle.read(4)
        for prefix, name in cls.BOMS:
            if head.startswith(prefix):
                return name

        with open(path, 'rb') as handle:
            raw = handle.read()
        try:
            raw.decode('utf-8')
        except UnicodeDecodeError:
            return cls.codepage()
        return 'utf-8'

    @classmethod
    def read(cls, key, value):
        """The value's data, or None when the command said it is not there.

        Kept public rather than name-mangled: the tests read through it, and
        Ruby's `private_class_method` has no enforcement-free Python twin
        that still reads as "not part of the surface".
        """
        out, ok = cls.run_reg(['query', key, '/v', value])
        if not ok:
            return None

        # "    NAME    TYPE    VALUE", indented, and wine leaves a CR on the
        # end. Stripping each line before splitting is what makes both go
        # away: without it the leading spaces produce an empty first field
        # and the value comes back as "REG_SZ    932\r\n".
        #
        # The name is matched as a whole field, not as a substring of the
        # line -- otherwise a value name that happens to appear inside
        # another value's data would match, and it would happen silently.
        matches = []
        for line in out.splitlines():
            parts = line.strip().split(None, 2)
            if len(parts) == 3 and parts[0] == value:
                matches.append(parts)

        if len(matches) > 1:
            raise VBAError(
                f"found {len(matches)} lines naming {value} in `wine reg query "
                f"{key}` output, expected exactly one: {out!r}")

        if not matches:
            raise VBAError(
                f"the command succeeded but no line naming {value} could be "
                f"parsed from `wine reg query {key}` output: {out!r}")

        return matches[0][2]

    @classmethod
    def write(cls, key, value, data):
        return cls.run_reg(
            ['add', key, '/v', value, '/t', 'REG_DWORD', '/d', data, '/f'])[1]

    @classmethod
    def run_reg(cls, args):
        """The only place that shells out, and deliberately not private: it
        is the seam the tests replace so that no test touches a real
        registry.

        Exit status is the only trustworthy signal. `reg` writes both its
        success message and its not-found message to stdout, in the system
        language, and wine writes unrelated `fixme:` lines to stderr.
        """
        try:
            proc = subprocess.run(
                ['wine', 'reg', *args],
                stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, check=False,
            )
        except OSError:
            # No wine on PATH is a FileNotFoundError, which is an OSError,
            # and so is every other way spawning can fail here.
            return ('', False)
        return (proc.stdout.decode('utf-8', 'replace'), proc.returncode == 0)
