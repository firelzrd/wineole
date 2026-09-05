import contextlib
import re

import wineole

from .address import Address
from .book import Book
from .passthrough import Passthrough
from .range import Range
from .sheet import Sheet

_DIGITS = re.compile(r'\A\d+\Z')


def _is_cell_reference(key):
    return (isinstance(key, tuple) and len(key) == 2
            and all(isinstance(k, int) and not isinstance(k, bool) for k in key))


class Excel(Passthrough):
    """Wraps a COM Excel.Application: the lifecycle (create / connect /
    connect_or_create / run) and the entry point of the addressing DSL that
    Address / Sheet / Book / Range build on.

    A `version` attribute is deliberately not defined -- see the comment on
    `_version` in `__init__`.

    `run` is a CLASS method, and unlike Ruby it IS reachable through an
    instance: Python finds classmethods on instances too, so `xl.run(...)`
    is the context manager, not COM's Application.Run (the macro runner --
    real on a live Excel 11, where a raw `run` call answers
    DISP_E_EXCEPTION rather than "unknown name"). Reach the macro runner as
    `xl.Run('Macro')` (COM names are case-sensitive here) or
    `xl.ole.Run('Macro')`; `xl.run('Macro')` raises ValueError for the
    unknown mode, and `xl.run('create')` would start a second Excel.
    """

    APPLICATION = 'Excel.Application'

    # What the bridge runs on the way out of an auto-created instance, once
    # its last user releases the root: suppress prompts, then quit.
    # Declared here, once, and handed to the bridge at construction time
    # (via Client.create/connect/connect_or_create's `cleanup=` argument)
    # rather than run from Python -- the bridge is the only party that knows
    # when the LAST user of a shared root has let go of it, so it is the
    # only party that can decide whether these steps should run at all.
    CLEANUP_STEPS = {'steps': [['DisplayAlerts=', False], ['Quit']]}

    @classmethod
    def create(cls, client=None, convert_paths=True):
        client = cls._client_or_default(client)
        return cls(client.create(cls.APPLICATION, cleanup=cls.CLEANUP_STEPS),
                   client, convert_paths=convert_paths)

    @classmethod
    def connect(cls, client=None, convert_paths=True):
        """Declares the same steps `create` does. That is correct, not an
        oversight: the steps are a property of this instance, but they only
        ever RUN when the bridge's record is auto-created, which is false
        here unless connect_or_create's own fallback is what created it. A
        human's Excel reached via connect is never auto-created, so Quit
        never fires for it."""
        client = cls._client_or_default(client)
        return cls(client.connect(cls.APPLICATION, cleanup=cls.CLEANUP_STEPS),
                   client, convert_paths=convert_paths)

    @classmethod
    def connect_or_create(cls, client=None, convert_paths=True):
        client = cls._client_or_default(client)
        return cls(client.connect_or_create(cls.APPLICATION, cleanup=cls.CLEANUP_STEPS),
                   client, convert_paths=convert_paths)

    @staticmethod
    def _client_or_default(client):
        # Looked up through the module, at call time, rather than imported
        # by name at import time: a caller (or a test) that replaces
        # wineole.default_client must be the one this sees.
        return wineole.default_client() if client is None else client

    @classmethod
    def run(cls, mode='connect_or_create', *, client=None, convert_paths=True):
        """Runs a block with an Excel application, then releases it:

            with Excel.run('create') as xl:
                ...

        Deciding whether that release actually quits Excel is not this
        method's job. create/connect/connect_or_create declare CLEANUP_STEPS
        at construction time, and the bridge is the one that knows when the
        LAST user of an auto-created root lets go of it -- that is when it
        runs those steps. Attaching to an Excel somebody already had open
        (connect) is never auto-created, so those steps never fire for it;
        quitting it out from under them would throw away their unsaved work.

        The mode is validated here, before the context manager is even
        built, so a typo fails at the call rather than at the `with`, and
        never after a connection has been made.
        """
        factories = {'create': cls.create,
                     'connect': cls.connect,
                     'connect_or_create': cls.connect_or_create}
        if not isinstance(mode, str) or mode not in factories:
            raise ValueError(f"unknown mode {mode!r}")
        factory = factories[mode]

        @contextlib.contextmanager
        def _runner():
            xl = factory(client=client, convert_paths=convert_paths)
            try:
                yield xl
            finally:
                xl.ole_release()

        return _runner()

    def __init__(self, proxy, client, convert_paths=True):
        self._ole = proxy
        self._client = client
        self._convert_paths = convert_paths
        # COM resolves names case-insensitively, so `xl.Version()` already
        # reaches COM's own Version and returns e.g. "11.0" -- measured.
        # Defining a `version` member here would only shadow that with
        # something that does the same thing, so this class deliberately
        # does not have one. Captured once here, rather than read fresh per
        # lookup, so every Sheet and Book this object builds sees one stable
        # answer without a round trip each time.
        self._version = proxy.Version()

    def ole_release(self):
        """Release the underlying Application. On the last user of an
        auto-created instance this is what quits Excel (the bridge runs
        DisplayAlerts=False then Quit); on a connected instance it simply
        detaches. Public because `run` is not the only way to get one of
        these, and a caller managing the lifecycle by hand needs the same
        call available."""
        return self._ole.ole_release()

    def leave_open(self):
        """Keep this Excel running after this program leaves -- e.g. a
        report left on screen for a human. Revokes the bridge's permission
        to quit it. One-way; there is no re-arm."""
        return self._ole.ole_leave_open()

    def show(self):
        # COM's Application has no Show/Hide (measured: both answer
        # DISP_E_UNKNOWNNAME), so defining these shadows nothing.
        self._ole.Visible = True

    def hide(self):
        self._ole.Visible = False

    @contextlib.contextmanager
    def no_alert(self):
        """Suppresses Excel's save/overwrite/etc. modal prompts for the
        duration of the block, then restores whatever DisplayAlerts was set
        to beforehand -- not a hardcoded True. Reading the value first means
        an outer caller who had already turned it off keeps it off.

        The leading read is deliberately OUTSIDE the try: if reading the
        flag is itself what failed, there is no previous value, and a
        restore of None would set the flag to False (measured) for the rest
        of the session. A failing read raises straight to the caller with no
        restore attempted.
        """
        previous = self._ole.DisplayAlerts()
        try:
            self._ole.DisplayAlerts = False
            yield
        finally:
            self._restore('DisplayAlerts', previous)

    @contextlib.contextmanager
    def no_update(self):
        """Same restore-what-was-there discipline as `no_alert`, for screen
        redraws instead of alert dialogs."""
        previous = self._ole.ScreenUpdating()
        try:
            self._ole.ScreenUpdating = False
            yield
        finally:
            self._restore('ScreenUpdating', previous)

    def __getitem__(self, key):
        return self._lookup(key)

    def __setitem__(self, key, value):
        """Resolves the same way `__getitem__` does; raises unless that
        resolves all the way to a Range. `xl["Sheet1!"] = 0` silently
        filling a whole sheet is exactly the hazard this guards against -- a
        typo that stops at a sheet or a book is otherwise indistinguishable
        from that "fill everything" request."""
        target = self._lookup(key)
        if not isinstance(target, Range):
            raise ValueError(
                f"{key!r} does not resolve to a range -- assignment needs an address "
                'with an explicit range; a bare sheet or workbook lookup is get-only')
        target.write(value)

    def _restore(self, name, value):
        """Put a flag back, and never let failing to do so become the
        caller's problem.

        An exception raised inside a finally REPLACES whatever the block
        raised. A failed restore would therefore destroy the caller's own
        error and report the cleanup instead.

        And the ordinary reason a restore fails is that the block ended the
        very thing being restored: an Application that quits or disconnects
        mid-block leaves nothing to put DisplayAlerts back on. There is no
        state left to restore, and nothing worth telling anyone.

        Swallowed unconditionally rather than only for a "the object is
        gone" error: over this bridge that arrives as a RemoteError wrapping
        an HRESULT string (0x800706BE is a normal transient right after
        Quit), so telling the two apart means matching on those strings --
        brittle, and beside the point, since raising out of a finally is
        wrong even when the failure really is transient.
        """
        try:
            setattr(self._ole, name, value)
        except Exception:
            pass

    def _lookup(self, key):
        if _is_cell_reference(key):
            return self._cell_range(key[0], key[1])

        if not isinstance(key, str):
            raise TypeError(f"unsupported index {key!r}")

        addr = Address.parse(key, self._version)

        # Not an address at all -- treat the whole string as a raw worksheet
        # name. This is exactly why Address.parse returns None instead of
        # raising.
        if addr is None:
            return Sheet(self._ole.Worksheets().Item(key), self._version)

        return self._resolve(addr, key)

    def _cell_range(self, row, col):
        # xl[row, col]: Cells on the active sheet. Goes through the same
        # no-Select path _resolve uses (via the Sheet it builds), one
        # Application round trip for ActiveSheet plus whatever Sheet's own
        # subscript costs.
        sheet = Sheet(self._active_worksheet(self._ole, f"xl[{row}, {col}]"), self._version)
        return sheet[row, col]

    def _resolve(self, addr, text):
        """Resolves a parsed Address against this Application.

        Deliberately never calls Select. Reaching a range through
        Application.Range would address whatever sheet happens to be active;
        resolving against the worksheet object itself needs no active state,
        so a read here does not mutate the caller's selection as a side
        effect, and it costs one fewer round trip. Measured against a live
        Excel 11: book.Worksheets().Item(1).Range('A1').Value() left
        ActiveSheet exactly where it was, both before and after.
        """
        workbook_ole = None if addr.workbook is None else self._resolve_workbook(addr.workbook, text)

        worksheet_part = addr.worksheet
        # A bare range ("A1:B2", no "!") names no worksheet at all -- it
        # still needs one to be looked up against, so it implicitly means
        # the active sheet, the same way xl[row, col] does.
        if worksheet_part is None and addr.has_range:
            worksheet_part = ''

        if worksheet_part is not None:
            container = self._ole if workbook_ole is None else workbook_ole
            worksheet_ole = self._resolve_worksheet(worksheet_part, container, text)
            sheet = Sheet(worksheet_ole, self._version)
            return sheet[addr.range] if addr.has_range else sheet

        # The grammar only leaves worksheet_part None when addr is the bare
        # "[workbook]" form -- worksheet is None and has_range is False
        # together only there -- so workbook_ole is always set by this point.
        return Book(workbook_ole, self._client, self._version,
                    convert_paths=self._convert_paths)

    def _resolve_workbook(self, part, text):
        if part == '':
            return self._active_workbook(text)
        if part.lower() == ':new':
            return self._ole.Workbooks().Add()
        return self._ole.Workbooks().Item(part)

    def _resolve_worksheet(self, part, container, text):
        if part == '':
            return self._active_worksheet(container, text)
        lowered = part.lower()
        if lowered == ':new':
            worksheets = container.Worksheets()
            return worksheets.Add(After=worksheets.Item(worksheets.Count()))
        if lowered == ':first':
            return container.Worksheets().Item(1)
        if lowered == ':last':
            worksheets = container.Worksheets()
            return worksheets.Item(worksheets.Count())
        if _DIGITS.match(part):
            return container.Worksheets().Item(int(part))
        return container.Worksheets().Item(part)

    # ActiveWorkbook/ActiveSheet are None on a fresh Excel with nothing open
    # yet (measured: Workbooks.Count == 0 makes both nil). Without these
    # checks that turns into an AttributeError raised from deep inside this
    # class instead of something a caller can act on.
    #
    # RuntimeError rather than ValueError: the address string itself was
    # fine -- it is the application's current state (no open workbook or
    # sheet) that cannot satisfy it.

    def _active_workbook(self, text):
        workbook = self._ole.ActiveWorkbook()
        if workbook is None:
            raise RuntimeError(f"no active workbook -- {text!r} needs one open")
        return workbook

    def _active_worksheet(self, container, text):
        worksheet = container.ActiveSheet()
        if worksheet is None:
            raise RuntimeError(f"no active worksheet -- {text!r} needs one open")
        return worksheet
