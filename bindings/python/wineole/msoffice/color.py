import re

_HEX = re.compile(r'\A#?(?:[0-9A-Fa-f]{3}|[0-9A-Fa-f]{6})\Z')


class Color:
    """'#RRGGBB' <-> Excel's colour integer.

    Excel stores a colour as BGR, not RGB: measured by making Excel name
    the colour itself, `Interior.Color = 255` reports ColorIndex 3 (red)
    and `= 0xFF0000` reports ColorIndex 5 (blue). So the obvious
    `int('FF0000', 16)` produces blue, silently.

    Public, and returning a plain int, on purpose. The wrapper's own keys
    are not the only place a colour is written -- the passthrough reaches
    Interior.Color, Font.Color, Borders.Color, Tab.Color and anything else
    COM has, and that surface is unbounded by design. One function that
    returns an int covers all of it without the protocol or Proxy's encode
    knowing anything about colours.
    """

    @staticmethod
    def parse(value):
        """'#RRGGBB' | '#RGB' | [r, g, b] | (r, g, b) -> Excel's integer."""
        r, g, b = Color._rgb(value)
        return r | (g << 8) | (b << 16)

    @staticmethod
    def to_hex(value):
        """Excel's integer (or the float it actually hands back) -> '#RRGGBB'."""
        # bool is an int in Python, and True is not a colour. Refuse it
        # here rather than answer '#000001' with total confidence.
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise TypeError(f"expected a number from COM, got {value!r}")

        # Checked before converting: int(math.inf) raises OverflowError and
        # int(nan) raises ValueError with a message about nan rather than
        # about colours. The comparison answers False for nan, so both land
        # here instead.
        if not (0 <= value <= 0xFFFFFF):
            raise ValueError(
                f"expected a colour in 0..0xFFFFFF, got {value!r}. "
                "Excel's ColorIndex is a different property from Color -- its "
                'values (-4105 automatic, -4142 none) are not colours and cannot '
                'be converted here'
            )
        n = int(value)
        return '#%02X%02X%02X' % (n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF)

    @staticmethod
    def _rgb(value):
        if isinstance(value, str):
            return Color._from_hex(value)
        if isinstance(value, (list, tuple)):
            return Color._from_sequence(value)
        raise TypeError(f"expected '#RRGGBB', '#RGB' or [r, g, b], got {value!r}")

    @staticmethod
    def _from_hex(value):
        # Checked before parsing: int(s, 16) on garbage raises with a
        # message about base 16 rather than about colours, and a three-digit
        # shorthand has to be expanded before it can be split at all.
        if not _HEX.match(value):
            raise ValueError(f"expected '#RRGGBB' or '#RGB', got {value!r}")

        s = value[1:] if value.startswith('#') else value
        if len(s) == 3:
            s = ''.join(c * 2 for c in s)
        return [int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)]

    @staticmethod
    def _from_sequence(value):
        # One message for every way the sequence can be wrong -- a wrong
        # length, a wrong element type and an out-of-range element all leave
        # the caller needing the same sentence. bool is excluded because it
        # is an int in Python and True is not a colour channel.
        ok = len(value) == 3 and all(
            isinstance(c, int) and not isinstance(c, bool) and 0 <= c <= 255
            for c in value
        )
        if not ok:
            raise ValueError(
                f"expected [r, g, b] with three integers in 0..255, got {value!r}")
        return list(value)
