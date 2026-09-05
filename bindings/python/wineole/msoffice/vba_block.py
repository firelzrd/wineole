"""A named span of code inside a VBA module, delimited by comment markers.

The wrapper owns the span, never the module. A module may be one the caller
wrote in, or one that cannot be deleted at all (ThisWorkbook, a worksheet),
so "replace the module" is not available and would be too blunt even where
it is.
"""

import re

from .vba import VBA


def split_com_lines(text):
    """Split a CodeModule's text on newlines, dropping trailing empties.

    Ruby's String#split drops every trailing empty field; Python's re.split
    keeps them. Everything below indexes into these lines and hands the
    leftovers to a caller that compares them, so the trailing '' that
    AddFromString's own final newline would produce has to go -- otherwise
    a removed block leaves a phantom blank line in the returned list.
    """
    parts = re.split(r'\r?\n', text)
    while parts and parts[-1] == '':
        parts.pop()
    return parts


def _chomp(text):
    """Ruby's String#chomp: drop ONE trailing line terminator, so a payload
    that ends in a blank line keeps that blank line before the close
    marker, exactly as the Ruby side writes it."""
    for end in ('\r\n', '\n', '\r'):
        if text.endswith(end):
            return text[:-len(end)]
    return text


class VBABlock:
    """A class used only through its classmethods; never instantiated."""

    NAME = re.compile(r'\A[A-Za-z0-9_-]+\Z')

    # Matches any wineole marker line, open or close, for any name -- not
    # just the one being written. Used to refuse a payload that would be
    # indistinguishable from the wrapper's own delimiters.
    MARKER_LINE = re.compile(r"\A'</?wineole:[A-Za-z0-9_-]+>\Z")

    @staticmethod
    def open_marker(name):
        return "'<wineole:%s>" % name

    @staticmethod
    def close_marker(name):
        return "'</wineole:%s>" % name

    @classmethod
    def write(cls, code_module, name, code):
        """Replaces the block of this name if it is there, adds it if not.

        A refused write (any of the three checks) never reaches the remove
        or the AddFromString, so nothing is touched.
        """
        cls._check_name(name)
        cls._check_representable(code)
        cls._check_payload(code)
        cls.remove(code_module, name)
        code_module.AddFromString(
            cls.open_marker(name) + '\n'
            + _chomp(code) + '\n'
            + cls.close_marker(name) + '\n')
        return None

    @classmethod
    def remove(cls, code_module, name):
        """False when there was nothing of this name to remove. Otherwise
        the module's remaining lines, as a list -- so a caller that needs to
        know whether the module is now blank does not have to fetch the
        whole body a second time to ask.

        Block names match case-insensitively, as VBA identifiers do: two
        blocks named `main` and `Main` would hold procedures that collide,
        and VBA answers "Ambiguous name detected" for the whole module from
        then on. So `Main` replaces (and removes) a block written as `main`.
        """
        cls._check_name(name)
        lines = cls._body(code_module)
        if not lines:
            return False

        open_marker = cls.open_marker(name).lower()
        close_marker = cls.close_marker(name).lower()
        first = None
        last = None
        for index, line in enumerate(lines):
            stripped = line.strip().lower()
            if first is None and stripped == open_marker:
                first = index
            if last is None and stripped == close_marker:
                last = index

        if first is None:
            if last is not None:
                raise ValueError(
                    f"the {name!r} block in this module has a closing marker "
                    'with no matching opening one -- the module is already '
                    'corrupted; refusing to guess what to remove')
            return False

        if last is None or last < first:
            raise ValueError(
                f"the {name!r} block in this module has no closing marker -- "
                'refusing to guess where it ends')

        code_module.DeleteLines(first + 1, last - first + 1)
        return lines[:first] + lines[last + 1:]

    @classmethod
    def blank(cls, code_module):
        """Nothing but whitespace. Not CountOfLines == 0: a module emptied
        of its blocks still reports the newlines that held them."""
        return cls.blank_lines(cls._body(code_module))

    @staticmethod
    def blank_lines(lines):
        """The same emptiness rule `blank` uses, applied to lines the caller
        already has (typically the list `remove` just handed back) instead
        of fetching the body again."""
        return all(line.strip() == '' for line in lines)

    @staticmethod
    def _body(code_module):
        """One round trip, whatever the module's length. Excel reports 0
        lines for a module never written to, and Lines(1, 0) is not a legal
        call."""
        count = int(code_module.CountOfLines())
        if count == 0:
            return []
        return split_com_lines(code_module.Lines(1, count))

    @classmethod
    def _check_representable(cls, code):
        """A module's text is held in the system ANSI codepage, not Unicode
        -- measured, not assumed: on a CP932 host `caf` + e-acute comes back
        `cafe`, the check mark comes back `?`, and simplified Chinese comes
        back part `?`. Japanese survives only because CP932 can represent
        it, which is why an earlier measurement using Japanese alone
        concluded, wrongly, that this path carried Unicode.

        So this path is bound by exactly the same codepage as import, and
        gets the same rule: refuse rather than substitute.

        ASCII skips the check, which is almost every call -- resolving the
        codepage costs a `wine reg` invocation.
        """
        if code.isascii():
            return
        try:
            code.encode(VBA.codepage())
        except UnicodeEncodeError as error:
            VBA.unrepresentable(error.object[error.start], 'this code')

    @classmethod
    def _check_name(cls, name):
        if not isinstance(name, str):
            raise TypeError(
                f"a block name must be a str, got {type(name).__name__}: {name!r}")
        if cls.NAME.fullmatch(name):
            return
        raise ValueError(
            f"a block name must match {cls.NAME.pattern!r} -- it goes inside a "
            "VBA comment marker, so it cannot contain spaces, '>' or newlines. "
            f"Got {name!r}")

    @classmethod
    def _check_payload(cls, code):
        """A code body containing a line that is itself a wineole marker
        would be indistinguishable from a real one once written: `remove`
        would find the caller's accidental marker instead of its own, delete
        up to the wrong place, and leave the rest as permanent garbage. The
        wrapper controls what it writes, so refusing the payload up front is
        what keeps that state from ever existing.
        """
        for line in re.split(r'\r?\n', code):
            stripped = line.strip()
            if cls.MARKER_LINE.fullmatch(stripped):
                raise ValueError(
                    'the code being written contains a line that is itself a '
                    f"wineole marker ({stripped!r}) -- this cannot be told apart "
                    "from the wrapper's own markers, so it is refused")
