from .color import Color

# Every number here was measured against a live Excel 11 rather than
# recalled.
UNDERLINE = {'none': -4142, 'single': 2, 'double': -4119}
ALIGN = {'general': 1, 'left': -4131, 'center': -4108, 'right': -4152, 'justify': -4130}
VALIGN = {'top': -4160, 'center': -4108, 'bottom': -4107}

XL_NONE = -4142
XL_AUTOMATIC = -4105
XL_GENERAL_FORMAT_NAME = 26  # Application.International(26)

# One knob for the caller, two properties for Excel. Excel keeps the line
# pattern (LineStyle) and its thickness (Weight) apart, which means 'thin'
# and 'dash' live in different properties even though a caller thinks of
# both as "what the line looks like". Measured pairs, (line, weight):
BORDER_STYLE = {
    'none': (-4142, None),
    'hairline': (1, 1),
    'thin': (1, 2),
    'medium': (1, -4138),
    'thick': (1, 4),
    'dash': (-4115, 2),
    'dot': (-4118, 2),
}

EDGE = {'left': 7, 'top': 8, 'bottom': 9, 'right': 10, 'inside_v': 11, 'inside_h': 12}
OUTLINE_EDGES = ('left', 'top', 'bottom', 'right')
ALL_EDGES = OUTLINE_EDGES + ('inside_v', 'inside_h')
BORDER_DICT_KEYS = ('edges', 'style', 'color')
ALL_EDGE_INDEXES = sorted(EDGE[name] for name in ALL_EDGES)

# Excel's own font size range. Outside it, COM fails with a message about
# the Font class rather than about the number.
SIZE_RANGE = range(1, 410)

FONT_KEYS = ('bold', 'italic', 'underline', 'size', 'color')
INTERIOR_KEYS = ('background',)
RANGE_KEYS = ('align', 'valign', 'wrap', 'number_format')
KEYS = FONT_KEYS + INTERIOR_KEYS + RANGE_KEYS + ('border',)


class Format:
    """The one place that knows how a format key maps onto a COM property.

    Keyword arguments rather than a chain, because formatting needs three
    states and a chain has two: an absent key means "leave it alone", and
    `False` means "explicitly turn it off". Those are different operations
    and both are ordinary. The implementation therefore asks `in opts` and
    never leans on `opts.get('bold')` being None.
    """

    @staticmethod
    def apply(ole, **opts):
        """Two passes on purpose. Everything is validated and converted into
        the values COM wants BEFORE the first write, so a bad key or a bad
        value leaves the range exactly as it was. Validating as it went
        would mean `format(bold=True, align='middle')` raises with the range
        already bold -- and a caller who sees an exception reasonably reads
        it as "nothing happened".

        `_translate` may READ from COM ('general' asks the Application for
        the local name of the General format). A read changes nothing, so
        the invariant that matters -- no write before validation finishes --
        still holds.
        """
        Format._reject_unknown_keys(opts)
        opts = {k: v for k, v in opts.items() if v is not None}  # None means "not specified"
        font, interior, range_level, border = Format._translate(ole, opts)

        # Each ole.Font() is its own round trip, so it is fetched once and
        # only when there is something to write to it.
        if font:
            Format._write_to(ole.Font(), font)
        if interior:
            Format._write_to(ole.Interior(), interior)
        Format._write_to(ole, range_level)
        if border is not None:
            Format._write_border(ole, border)
        return None

    @staticmethod
    def _write_to(target, assignments):
        for name, value in assignments:
            setattr(target, name, value)

    @staticmethod
    def _translate(ole, opts):
        """Returns four values: three lists of (property name, value) -- for
        Font, for Interior and for the Range itself -- plus a border plan
        (or None) for `_write_border`, which does not fit the two-tuple
        shape because it may need both a bulk assignment and a per-edge
        loop. Raises rather than returning anything partial."""
        font = []
        interior = []
        range_level = []

        if 'bold' in opts:
            font.append(('Bold', Format._boolean(opts['bold'], 'bold')))
        if 'italic' in opts:
            font.append(('Italic', Format._boolean(opts['italic'], 'italic')))
        if 'underline' in opts:
            font.append(('Underline', Format._underline(opts['underline'])))
        if 'size' in opts:
            font.append(('Size', Format._size(opts['size'])))
        if 'color' in opts:
            if opts['color'] is False:
                font.append(('ColorIndex', XL_AUTOMATIC))
            else:
                font.append(('Color', Format._colour(opts['color'], 'color')))

        if 'background' in opts:
            if opts['background'] is False:
                # Not `Color = white`: measured, a cleared cell and a
                # white-painted cell both report Color 16777215, but the
                # painted one keeps ColorIndex 2 and Pattern 1 -- it is
                # still filled, prints as a fill, and hides gridlines.
                interior.append(('ColorIndex', XL_NONE))
            else:
                interior.append(('Color', Format._colour(opts['background'], 'background')))

        if 'align' in opts:
            range_level.append(('HorizontalAlignment', Format._fetch(ALIGN, opts['align'], 'align')))
        if 'valign' in opts:
            range_level.append(('VerticalAlignment', Format._fetch(VALIGN, opts['valign'], 'valign')))
        if 'wrap' in opts:
            range_level.append(('WrapText', Format._boolean(opts['wrap'], 'wrap')))
        if 'number_format' in opts:
            range_level.append(('NumberFormat', Format._number_format(ole, opts['number_format'])))

        border = Format._translate_border(opts['border']) if 'border' in opts else None
        return font, interior, range_level, border

    @staticmethod
    def _translate_border(spec):
        """Validates and resolves everything; touches no COM."""
        spec = Format._normalize_border(spec)
        line, weight = Format._fetch(BORDER_STYLE, spec['style'], 'border style')
        line_colour = Format._colour(spec['color'], 'border color') if 'color' in spec else None
        return {'indexes': Format._expand_edges(spec.get('edges')),
                'line': line, 'weight': weight, 'colour': line_colour}

    @staticmethod
    def _write_border(ole, plan):
        """ole.Borders() is a round trip like Font and Interior: fetched
        once, with Item() called off it -- except when every edge is being
        set, in which case an assignment straight to the Borders collection
        replaces the whole per-edge loop.

        Measured against a live Excel on a multi-cell range, an assignment
        on Borders itself reaches all six edges -- including inside_v and
        inside_h -- in one COM call each: LineStyle, Weight and Color each
        touch all six for the price of one round trip. 2.3 ms against
        10.0 ms for the per-edge Item() loop below. That is exactly why it
        must never be used for 'outline': it would silently draw the inside
        edges too. Keyed off the resolved index set rather than the 'all'
        name, so an explicit list of all six edges gets the fast path too.
        """
        indexes = list(dict.fromkeys(plan['indexes']))  # uniq, order preserved
        if not indexes:
            return

        borders = ole.Borders()

        if sorted(indexes) == ALL_EDGE_INDEXES:
            borders.LineStyle = plan['line']
            if plan['weight'] is None:
                return
            borders.Weight = plan['weight']
            if plan['colour'] is not None:
                borders.Color = plan['colour']
            return

        for index in indexes:
            edge = borders.Item(index)
            edge.LineStyle = plan['line']
            # Nothing to weigh or colour when the line is being removed.
            if plan['weight'] is None:
                continue
            edge.Weight = plan['weight']
            if plan['colour'] is not None:
                edge.Color = plan['colour']

    @staticmethod
    def _normalize_border(spec):
        if spec is False:
            out = {'edges': 'all', 'style': 'none'}
        elif isinstance(spec, str):
            out = {'edges': spec, 'style': 'thin'}
        elif isinstance(spec, (list, tuple)):
            out = {'edges': spec, 'style': 'thin'}
        elif isinstance(spec, dict):
            out = spec
        else:
            raise ValueError(
                "border: expected 'all', 'outline', an edge name, a list of edge "
                f"names, False, or a dict, got {spec!r}")

        # Checked against the dict as given, before Nones are dropped --
        # same order as apply's top-level keys, so a misspelled key with a
        # None value is still caught rather than silently absorbed.
        unknown = [k for k in out if k not in BORDER_DICT_KEYS]
        if unknown:
            plural = 's' if len(unknown) > 1 else ''
            names = ', '.join(repr(k) for k in unknown)
            raise ValueError(
                f"border: unknown key{plural} {names} -- known keys are "
                + ', '.join(BORDER_DICT_KEYS))

        # None means "not specified" everywhere else in this module (apply
        # drops None top-level values before anything is validated); a
        # caller writing style=None explicitly is asking for the same thing
        # as omitting style, not for an override that beats the default.
        merged = {'style': 'thin'}
        merged.update({k: v for k, v in out.items() if v is not None})
        return merged

    @staticmethod
    def _expand_edges(edges):
        if edges is None:
            # Only a dict form reaches here without edges having been filled
            # in -- the shorthands ('all', an edge name, a list, False) all
            # set it themselves in _normalize_border. Naming the missing key
            # beats _bad_edges_message's "got None", which reads as a value
            # the caller never wrote.
            raise ValueError(
                "border: a dict needs an edges key, e.g. "
                "border={'edges': 'all', 'style': 'thick'}")
        if edges == 'all':
            names = list(ALL_EDGES)
        elif edges == 'outline':
            names = list(OUTLINE_EDGES)
        elif isinstance(edges, str):
            names = [edges]
        elif isinstance(edges, (list, tuple)):
            # 'all' and 'outline' expand here too, not just on their own --
            # the error message below promises "a list of those", and
            # "those" includes the shorthands. This also makes something
            # like ['outline', 'inside_h'] expressible.
            names = []
            for edge in edges:
                if edge == 'all':
                    names.extend(ALL_EDGES)
                elif edge == 'outline':
                    names.extend(OUTLINE_EDGES)
                else:
                    names.append(edge)
        else:
            raise ValueError(Format._bad_edges_message(edges))

        out = []
        for name in names:
            if not isinstance(name, str) or name not in EDGE:
                raise ValueError(Format._bad_edges_message(name))
            out.append(EDGE[name])
        return out

    @staticmethod
    def _bad_edges_message(value):
        # One message for both failures, and it names the shorthands as well
        # as the edges -- someone who typed 'diagonal' needs to learn that
        # 'all' and 'outline' exist, which a bare list of EDGE's keys would
        # not tell them.
        return ("border: expected 'all', 'outline', one of "
                + ', '.join(repr(k) for k in EDGE)
                + f", or a list of those, got {value!r}")

    @staticmethod
    def _reject_unknown_keys(opts):
        # Before any COM call, so a typo leaves the sheet exactly as it was
        # rather than half-formatted.
        unknown = [k for k in opts if k not in KEYS]
        if not unknown:
            return
        plural = 's' if len(unknown) > 1 else ''
        names = ', '.join(repr(k) for k in unknown)
        raise ValueError(
            f"unknown format key{plural} {names} -- known keys are " + ', '.join(KEYS))

    @staticmethod
    def _boolean(value, key):
        # A None never reaches here -- apply drops None values first,
        # because None means "I have no value for this", which is the same
        # thing as not passing the key. That matters: measured, assigning
        # nil to a COM boolean property sets it to FALSE, so a None reaching
        # COM would silently un-bold a range whose caller simply did not know.
        if value is True or value is False:
            return value
        raise TypeError(
            f"{key}: expected True or False, got {value!r}. "
            'Omit the key entirely to leave this attribute alone')

    @staticmethod
    def _size(value):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise TypeError(
                f"size: expected a number in 1..409 (Excel's own range), got {value!r}")
        # Not `value in SIZE_RANGE`: a range only contains integers, and
        # Excel accepts 14.5. The endpoints are what the check needs.
        if not (SIZE_RANGE.start <= value <= SIZE_RANGE.stop - 1):
            raise ValueError(
                f"size: expected a number in 1..409 (Excel's own range), got {value!r}")
        return value

    @staticmethod
    def _underline(value):
        if value is True:
            return UNDERLINE['single']
        if value is False:
            return UNDERLINE['none']
        return Format._fetch(UNDERLINE, value, 'underline')

    @staticmethod
    def _colour(value, key):
        # A raw COM colour integer is ambiguous here: 255 could mean the
        # caller's #0000FF or Excel's own value for red. Refuse rather than
        # guess -- the same stance `write` takes on a wrong-shaped list.
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            raise ValueError(
                f"{key}: expected '#RRGGBB', got the number {value!r}. "
                'A raw COM colour is ambiguous here -- pass a hex string, or use '
                'wineole.msoffice.Color.parse(...) with .ole to reach COM directly')
        try:
            return Color.parse(value)
        except ValueError as err:
            raise ValueError(f"{key}: {err}") from None
        except TypeError as err:
            raise TypeError(f"{key}: {err}") from None

    @staticmethod
    def _number_format(ole, value):
        # 'General' is the one format code that cannot be written:
        # measured, it fails outright on a localized Excel, where the format
        # has a translated name instead -- and that translated spelling is
        # not portable either. Application.International(26) returns
        # whichever one this Excel wants. Note the () on Application(): the
        # Python raw client needs it.
        if not isinstance(value, str):
            raise TypeError(
                "number_format: expected a format code string, 'general' or "
                f"'text', got {value!r}")
        if value.lower() == 'general':
            return ole.Application().International(XL_GENERAL_FORMAT_NAME)
        if value.lower() == 'text':
            return '@'
        return value

    @staticmethod
    def _fetch(table, value, key):
        try:
            return table[value]
        except (KeyError, TypeError):
            raise ValueError(
                f"{key}: expected one of " + ', '.join(repr(k) for k in table)
                + f", got {value!r}") from None
