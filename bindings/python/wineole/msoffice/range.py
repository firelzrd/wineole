from .format import Format
from .passthrough import Passthrough


class Range(Passthrough):
    """Wraps a COM Range. Adds exactly four methods plus `ole`; everything
    else falls through to COM.

    The restraint is deliberate. COM's Range already exposes Rows, Columns,
    Cells, Item, Areas, Find, Sort, Merge, Table and many more, and COM
    resolves names case-insensitively -- so every lowercase single word this
    class defines covers a COM call that already worked. `to_list`, `write`,
    `fill`, `format` and `ole` were each checked against a live Range and
    found absent.
    """

    def __init__(self, proxy):
        self._ole = proxy

    def to_list(self):
        """Always a two-dimensional list, whatever the range's size.

        Excel's own `Value` returns a bare scalar for a one-cell range --
        consistently, whether it was addressed as "A1" or "A1:A1" -- so
        generic code that does not know the size has to branch. This is that
        branch, written once.

        Returns a plain list on purpose: Python's own vocabulary then covers
        the rest (`list(zip(*rows))` for columns, a comprehension for
        values), so this class does not have to grow `each_row`,
        `each_column` or `flatten` and risk shadowing more COM members.

        Named `to_list` rather than Ruby's `to_a`: `to_a` means nothing in
        Python, and `__iter__` is deliberately refused (see Passthrough), so
        the conversion has to be an ordinary explicit method.
        """
        v = self._ole.Value()
        return v if isinstance(v, list) else [[v]]

    def write(self, value):
        """Write, refusing anything that does not fit.

        Excel's own assignment corrupts silently in three ways (all
        measured): a flat list written to a column replicates its first
        element down every cell, too few values leave #N/A behind, and too
        many are truncated. None of those raise, and none are visible
        without looking at the sheet.
        """
        self._ole.Value = self._shaped(value)

    def fill(self, value):
        """Write, adapting the value to the range: replicate along a
        dimension the argument does not have, truncate or pad along one it
        does.

        Total -- every input has a defined result, no exceptions -- but note
        that it reproduces Excel's own column trap by construction: a flat
        list is a row, so filling an Nx1 column with [1, 2, 3] puts 1 in
        every cell. That is why `write` and not `fill` is what
        `sheet[addr] = x` uses; reach for this one deliberately.
        """
        nrows = self._row_count()
        ncols = self._column_count()
        if isinstance(value, (list, tuple)):
            rows = list(value)
            if rows and isinstance(rows[0], (list, tuple)):
                out = []
                for r in range(nrows):
                    row = list(rows[r]) if r < len(rows) else []
                    out.append([row[c] if c < len(row) else None for c in range(ncols)])
            else:
                one = [rows[c] if c < len(rows) else None for c in range(ncols)]
                out = [list(one) for _ in range(nrows)]
        else:
            out = [[value for _ in range(ncols)] for _ in range(nrows)]
        self._ole.Value = out

    def format(self, **opts):
        """Apply formatting. Keys are documented on
        `wineole.msoffice.format.Format`.

        An absent key leaves that attribute alone; `False` turns it off.
        That third state is why this takes keyword arguments rather than
        being a chain of verbs -- and it keeps this class's additions to one
        name, which matters because COM resolves names case-insensitively
        and every lowercase word here covers a COM member of the same
        spelling. `format` was measured free on a live Range.
        """
        Format.apply(self._ole, **opts)
        return self

    def _row_count(self):
        return self._ole.Rows().Count()

    def _column_count(self):
        return self._ole.Columns().Count()

    def _shaped(self, value):
        # A str is iterable in Python but is a scalar here, so only list and
        # tuple count as "rows or values"; everything else goes through
        # untouched for Excel to broadcast itself.
        if not isinstance(value, (list, tuple)):
            return value

        rows = list(value)
        nrows = self._row_count()
        ncols = self._column_count()

        if rows and isinstance(rows[0], (list, tuple)):
            if not all(isinstance(r, (list, tuple)) for r in rows):
                raise ValueError(
                    f"range is {nrows}x{ncols}, but the value mixes rows and scalars")
            widths = []
            for r in rows:
                if len(r) not in widths:
                    widths.append(len(r))
            if len(widths) > 1:
                raise ValueError(
                    f"range is {nrows}x{ncols}, but the value has ragged rows "
                    f"({', '.join(str(w) for w in widths)} elements)")
            if any(isinstance(c, (list, tuple)) for r in rows for c in r):
                raise ValueError(
                    f"range is {nrows}x{ncols}, but the value nests more than two deep")
            if not (len(rows) == nrows and widths[0] == ncols):
                raise ValueError(
                    f"range is {nrows}x{ncols}, but the value is {len(rows)}x{widths[0]}")
            return [list(r) for r in rows]

        if any(isinstance(v, (list, tuple)) for v in rows):
            raise ValueError(
                f"range is {nrows}x{ncols}, but the value mixes scalars and rows")
        if nrows > 1 and ncols > 1:
            raise ValueError(
                f"range is {nrows}x{ncols}; a flat list only fits a single row or "
                'column -- pass rows, or use fill')
        expected = nrows if nrows > 1 else ncols
        if len(rows) != expected:
            raise ValueError(
                f"range is {nrows}x{ncols}, but the value has {len(rows)} elements")
        return [[v] for v in rows] if nrows > 1 else [rows]
