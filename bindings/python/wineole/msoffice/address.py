import re


def _group(pattern):
    """Wrap a fragment so concatenating it cannot leak an alternation.

    Ruby interpolates a Regexp into another one as `(?-mix:...)`, which
    groups it. Python's f-string interpolation does not, so
    `A|B` + `C|D` would parse as `A | BC | D`. Every fragment below is
    therefore wrapped here, once, rather than by hand at each use.
    """
    return '(?:' + pattern + ')'


# Excel 11's grid: columns A..IV, rows 1..65536.
_A_IV = _group('(?i:[A-H]?[A-Z]|I[A-V])')
_ABS_A_IV = _group(r'\$?' + _A_IV)
_1_65536 = _group(r'[1-9]\d{0,3}|[1-5]\d{4}|6(?:[0-4]\d{3}|5(?:[0-4]\d{2}|5(?:[0-2]\d|3[0-6])))')
_ABS_1_65536 = _group(r'\$?' + _1_65536)
_ABS_A1_IV65536 = _group(_ABS_A_IV + _ABS_1_65536)
_RANGE_XL11 = (
    '(?P<range>'
    + _group(_ABS_A1_IV65536 + ':') + '?' + _ABS_A1_IV65536
    + '|' + _ABS_A_IV + ':' + _ABS_A_IV
    + '|' + _ABS_1_65536 + ':' + _ABS_1_65536
    + ')'
)

# Excel 12's grid: columns A..XFD, rows 1..1048576.
_A_XFD = _group('(?i:[A-W]?[A-Z]{1,2}|X(?:[A-E][A-Z]|F[A-D]))')
_ABS_A_XFD = _group(r'\$?' + _A_XFD)
_1_1048576 = _group(
    r'[1-9]\d{0,5}|10(?:[0-3]\d{4}|4(?:[0-7]\d{3}|8(?:[0-4]\d{2}|5(?:[0-6]\d|7[0-6]))))'
)
_ABS_1_1048576 = _group(r'\$?' + _1_1048576)
_ABS_A1_XFD1048576 = _group(_ABS_A_XFD + _ABS_1_1048576)
_RANGE_XL12 = (
    '(?P<range>'
    + _group(_ABS_A1_XFD1048576 + ':') + '?' + _ABS_A1_XFD1048576
    + '|' + _ABS_A_XFD + ':' + _ABS_A_XFD
    + '|' + _ABS_1_1048576 + ':' + _ABS_1_1048576
    + ')'
)

_WORKBOOK = r'\[(?P<workbook>(?i::new)|[^\[\]]*)\]'
_WORKBOOK_WORKSHEET = (
    "(?P<worksheet_quote>'?)"
    + _group(_WORKBOOK) + '?'
    + "(?P<worksheet>(?i::(?:new|first|last))|(?:[^\\[\\]\\\\:\\*'][^\\[\\]\\\\:\\*]*)?)"
    + '(?P=worksheet_quote)!'
)


def _compile(range_pattern):
    return re.compile(r'\A' + _group(_WORKBOOK_WORKSHEET) + '?' + _group(range_pattern) + r'?\Z')


class Address:
    """Parses the addressing DSL:

        "[Book1]Sheet1!A1:B2"   "[:new]"   ":new!"   ":first!A1"   "A1:B2"

    Touches no COM: a pure string parser, so it can be exercised without
    Excel running.

    Two anchored patterns rather than Ruby's one, and that is forced:
    Ruby's `/\\A(?:WORKBOOK|WORKBOOK_WORKSHEET?RANGE?)\\z/` names the
    `workbook` group in both branches, and Python's `re` refuses a group
    name that appears twice in one pattern. The book-only pattern is
    therefore tried first and the full one second -- the same order Ruby's
    alternation tries them in, so the acceptance set and the captures are
    identical. The address test file proves it with every form of the Ruby
    test file plus the grid boundaries.
    """

    PTN_BOOK_ONLY = re.compile(r'\A' + _WORKBOOK + r'\Z')
    PTN_XL11 = _compile(_RANGE_XL11)
    PTN_XL12 = _compile(_RANGE_XL12)

    def __init__(self, workbook, worksheet, range):
        self.workbook = workbook
        self.worksheet = worksheet
        self.range = range

    @staticmethod
    def parse(text, excel_version):
        """Returns None when the string is not an address at all, so a
        caller can fall back to treating it as a raw sheet name -- which is
        exactly what Excel._lookup does."""
        if not isinstance(text, str):
            return None

        m = Address.PTN_BOOK_ONLY.match(text)
        if m is not None:
            return Address(m.group('workbook'), None, None)

        pattern = Address.PTN_XL12 if Address._is_xl12(excel_version) else Address.PTN_XL11
        m = pattern.match(text)
        # An empty match means the pattern's every part was optional and
        # every one of them was absent -- "" is not an address.
        if m is None or m.group(0) == '':
            return None
        return Address(m.group('workbook'), m.group('worksheet'), m.group('range'))

    @staticmethod
    def _is_xl12(excel_version):
        # Ruby reads the version with String#to_f, which answers 0.0 for
        # anything unparseable rather than raising. Same leniency here: a
        # version this code cannot read is treated as the smaller grid,
        # which refuses more addresses rather than accepting ones the host
        # cannot hold.
        try:
            return float(excel_version) >= 12
        except (TypeError, ValueError):
            return False

    @property
    def has_range(self):
        """Whether this address names a range of cells. An address that
        stops at a sheet or a book is a lookup, not an assignment target --
        `xl["Sheet1!"] = 0` filling a whole sheet is a hazard rather than a
        convenience."""
        return self.range is not None
