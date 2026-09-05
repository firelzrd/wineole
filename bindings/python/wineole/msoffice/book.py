from .forms import Forms
from .passthrough import Passthrough
from .paths import Paths
from .sheet import Sheet
from .vba_api import BookVBA


class Book(Passthrough):
    """Wraps a COM Workbook.

    Name note (measured against a live Excel 11): `sheet`, `save_as` and
    `local_path` are all free -- COM answers DISP_E_UNKNOWNNAME for each.
    `close` is not: COM resolves member names case-insensitively, and
    `Workbook.Close` already exists. It is a deliberate shadow anyway (see
    `close` below). `local_file` is free by the same construction as
    `save_as`: COM does not strip underscores when matching names, so it
    never collides with `Workbook.FullName`.
    """

    def __init__(self, proxy, client, version, convert_paths=True):
        # `convert_paths` is the caller's opt-out; whether converting is
        # even meaningful is a separate question (a remote bridge's own
        # filesystem means nothing to this machine). Combining them once
        # here means a caller cannot talk a remote bridge into converting by
        # passing convert_paths=True -- Paths.convertible still says no.
        self._ole = proxy
        self._version = version
        self._convert_paths = bool(convert_paths and Paths.convertible(client))
        # Underscore-prefixed, so Passthrough.__setattr__ keeps it here
        # rather than sending it to COM. Built on first use: reaching a
        # VBProject is a COM call, and most books never need one.
        self._vba = None
        # Same shape as _vba: underscored so Passthrough keeps it, and built
        # on first use because most books have no UserForms.
        self._forms = None

    def sheet(self, name_or_index):
        """Worksheets, not Sheets: Sheets also includes chart sheets, which
        this wrapper's Sheet class does not model."""
        return Sheet(self._ole.Worksheets().Item(name_or_index), self._version)

    def sheets(self):
        """A generator over this workbook's worksheets, in order.

        Ruby's `each_sheet` takes a block and answers an Enumerator without
        one; a generator is the single Python spelling that covers both --
        `for s in book.sheets():` and `list(book.sheets())` both work, and
        nothing is fetched until it is iterated.
        """
        worksheets = self._ole.Worksheets()
        for i in range(1, worksheets.Count() + 1):
            yield Sheet(worksheets.Item(i), self._version)

    def save_as(self, path):
        """Takes only the path. A caller needing FileFormat and the rest of
        COM's SaveAs arguments uses the passthrough `book.SaveAs(...)`."""
        target = Paths.to_wine(path) if self._convert_paths else path
        return self._ole.SaveAs(target)

    @property
    def local_path(self):
        """COM's Workbook.Path is the CONTAINING FOLDER, not the file --
        local_path deliberately names the same thing this wrapper's way, in
        Linux form. The file's own path is `local_file`.

        An unsaved book's Path is "" -- Paths.to_local returns that
        unchanged without shelling out to winepath.
        """
        value = self._ole.Path()
        return Paths.to_local(value) if self._convert_paths else value

    @property
    def local_file(self):
        """The file's own path, in Linux form -- what local_path is not.

        Gated by the same loopback-only rule as local_path; calling
        Paths.to_local(book.FullName()) directly instead skips that gate and
        runs a local winepath over what may be a REMOTE bridge's Wine path,
        silently producing a path that refers to a filesystem this machine
        does not have. A wrong conversion that happens silently is worse
        than one that visibly does not happen.

        Measured against a live Excel 11: FullName is the file
        (Z:\\tmp\\wineole_item_probe.xls where Path is the folder, Z:\\tmp);
        an unsaved book's FullName is the bare in-memory name ("Book1"),
        matching its Path of "".
        """
        value = self._ole.FullName()
        return Paths.to_local(value) if self._convert_paths else value

    def close(self, save=False):
        """A deliberate shadow of COM's Workbook.Close. Close with no
        SaveChanges argument can raise a modal save-changes prompt, which
        under Wine is a hang; close(save=False) turns that hazard into an
        explicit parameter instead. The raw member stays reachable as
        `book.Close(...)` (exact PascalCase) and as `book.ole.Close(...)`.
        """
        return self._ole.Close(save)

    @property
    def vba(self):
        """This workbook's VBA surface: blocks, components, import/export.

        A property, not a method: zero-argument accessors are properties in
        this package (`local_path`, `local_file`). Memoized, so `book.vba`
        twice is one object and the module it makes on demand is made once.
        """
        if self._vba is None:
            self._vba = BookVBA(self._ole, convert_paths=self._convert_paths)
        return self._vba

    @property
    def forms(self):
        """This workbook's UserForms. Memoised, so `book.forms` twice is one
        `Forms` object and the helper block has one owner.

        Name note (measured against a live Excel 11): `forms` is free on
        Workbook.
        """
        if self._forms is None:
            self._forms = Forms(self)
        return self._forms
