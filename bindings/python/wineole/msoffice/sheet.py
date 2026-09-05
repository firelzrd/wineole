from .address import Address
from .controls import ActiveXControls, FormControls
from .passthrough import Passthrough
from .range import Range
from .vba_api import SheetVBA


def _is_cell_reference(key):
    return (isinstance(key, tuple) and len(key) == 2
            and all(isinstance(k, int) and not isinstance(k, bool) for k in key))


def _named(part):
    return part is not None and part != ''


class Sheet(Passthrough):
    """Wraps a COM Worksheet. `[]`/`[]=` address it through the same
    `Address` parser the whole wrapper uses, so `sheet['A1:B2']` and
    `sheet[row, col]` both hand back a `Range`; everything else falls
    through to COM.
    """

    def __init__(self, proxy, version):
        self._ole = proxy
        self._version = version
        self._vba = None
        # Underscore-prefixed like every wrapper attribute, so
        # Passthrough.__setattr__ keeps them here rather than sending them to
        # COM. Built on first use, and memoised: a collection that was a
        # fresh object per access would hand `add` and the `[]` that finds
        # the result different BookVBA writers.
        self._form_controls = None
        self._activex = None

    def __getitem__(self, key):
        return self._range_for(key)

    def __setitem__(self, key, value):
        # Deliberately Range.write, never fill: write raises on a value that
        # does not fit the range rather than replicating or padding it,
        # which is the behaviour an assignment through []= should have.
        self._range_for(key).write(value)

    def _range_for(self, key):
        if _is_cell_reference(key):
            return Range(self._ole.Cells(key[0], key[1]))

        if not isinstance(key, str):
            raise TypeError(f"unsupported sheet index {key!r}")

        addr = Address.parse(key, self._version)

        # An address that parses but stops at a sheet or a book (or does not
        # parse at all) is a lookup, not something with cells to read or
        # write -- `xl["Sheet1!"] = 0` filling a whole sheet is a hazard
        # rather than a convenience.
        if addr is None or not addr.has_range:
            raise ValueError(f"{key!r} has no range")

        # This object is one sheet. An address that names a different
        # workbook or worksheet would silently reach past it -- writing to
        # the sheet the caller NAMED instead of the one they HAVE is exactly
        # the class of silent wrong-target write this wrapper exists to
        # prevent.
        if _named(addr.workbook) or _named(addr.worksheet):
            raise ValueError(
                f"{key!r} names another workbook or worksheet; a Sheet "
                'addresses its own cells only -- reach another sheet through '
                'the Excel object that owns it')

        return Range(self._ole.Range(addr.range))

    @property
    def vba(self):
        """This worksheet's VBA surface: blocks in the sheet's own code
        module, which is where Excel looks for an ActiveX control's
        `_Click`. Blocks only -- a worksheet's module can be neither created
        nor deleted."""
        if self._vba is None:
            self._vba = SheetVBA(self._ole)
        return self._vba

    @property
    def form_controls(self):
        """The Forms-toolbar controls on this sheet. Measured free on
        Worksheet: COM answers DISP_E_UNKNOWNNAME for this name."""
        if self._form_controls is None:
            self._form_controls = FormControls(self)
        return self._form_controls

    @property
    def activex(self):
        """The worksheet ActiveX controls on this sheet -- the family whose
        events reach Python. Measured free on Worksheet."""
        if self._activex is None:
            self._activex = ActiveXControls(self)
        return self._activex
