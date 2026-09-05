import math
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import FakeComRangeForFormat, LOCAL_GENERAL

from wineole.msoffice.color import Color
from wineole.msoffice.format import Format


class MSOfficeFormatTest(unittest.TestCase):
    def setUp(self):
        self.ole = FakeComRangeForFormat()

    def apply(self, **opts):
        Format.apply(self.ole, **opts)
        return self.ole

    # --- the three-state rule ------------------------------------------

    def test_an_absent_key_touches_nothing(self):
        self.apply(italic=True)
        self.assertNotIn('Bold', self.ole.font.writes,
                         'a key that was not passed must not be written at all -- '
                         'absent means "leave it alone"')

    def test_false_is_written_rather_than_skipped(self):
        self.apply(bold=False)
        self.assertIs(False, self.ole.font.writes['Bold'],
                      'False means "explicitly turn it off", which is not the same as '
                      'leaving it alone')

    def test_nothing_at_all_touches_nothing(self):
        self.apply()
        self.assertEqual({}, self.ole.writes)
        self.assertEqual(0, self.ole.font_fetches)
        self.assertEqual(0, self.ole.interior_fetches)

    # --- font ----------------------------------------------------------

    def test_bold_and_italic_and_size(self):
        self.apply(bold=True, italic=True, size=14)
        self.assertIs(True, self.ole.font.writes['Bold'])
        self.assertIs(True, self.ole.font.writes['Italic'])
        self.assertEqual(14, self.ole.font.writes['Size'])

    # Font.Underline is not a boolean: measured, its default is -4142 and
    # single underline is 2. Sending True would mean something else.
    def test_underline_true_becomes_single_not_true(self):
        self.apply(underline=True)
        self.assertEqual(2, self.ole.font.writes['Underline'])

    def test_underline_false_becomes_none(self):
        self.apply(underline=False)
        self.assertEqual(-4142, self.ole.font.writes['Underline'])

    def test_underline_names(self):
        for name, expected in (('single', 2), ('double', -4119), ('none', -4142)):
            self.setUp()
            self.assertEqual(expected, self.apply(underline=name).font.writes['Underline'], name)

    def test_an_unknown_underline_name_is_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(underline='wavy')
        self.assertIn('underline', str(ctx.exception))
        self.assertIn('single', str(ctx.exception))

    # --- colour --------------------------------------------------------

    def test_colour_goes_through_the_bgr_conversion(self):
        self.apply(color='#FF0000', background='#0000FF')
        self.assertEqual(255, self.ole.font.writes['Color'], 'red must be 255, not 0xFF0000')
        self.assertEqual(16711680, self.ole.interior.writes['Color'])

    # Color cannot express "no fill": a cleared cell and a white-painted
    # cell both read back 16777215. Clearing has to go through ColorIndex.
    def test_background_false_clears_rather_than_painting_white(self):
        self.apply(background=False)
        self.assertEqual(-4142, self.ole.interior.writes['ColorIndex'])
        self.assertNotIn('Color', self.ole.interior.writes,
                         'clearing must not write a Color at all -- painting white leaves '
                         'the cell filled')

    def test_color_false_sets_automatic(self):
        self.apply(color=False)
        self.assertEqual(-4105, self.ole.font.writes['ColorIndex'])
        self.assertNotIn('Color', self.ole.font.writes)

    def test_a_raw_com_colour_integer_is_refused(self):
        for raw in [255, 0xFF0000, 0]:
            with self.assertRaises(ValueError, msg=f"{raw} must be refused") as ctx:
                Format.apply(FakeComRangeForFormat(), color=raw)
            self.assertIn('Color', str(ctx.exception),
                          'the message must point at msoffice Color')

    # --- alignment, wrap ------------------------------------------------

    def test_alignment_names_become_the_measured_numbers(self):
        self.apply(align='center', valign='top', wrap=True)
        self.assertEqual(-4108, self.ole.writes['HorizontalAlignment'])
        self.assertEqual(-4160, self.ole.writes['VerticalAlignment'])
        self.assertIs(True, self.ole.writes['WrapText'])

    def test_every_alignment_name(self):
        for name, num in {'general': 1, 'left': -4131, 'center': -4108,
                          'right': -4152, 'justify': -4130}.items():
            self.setUp()
            self.assertEqual(num, self.apply(align=name).writes['HorizontalAlignment'],
                             f"align: {name}")
        for name, num in {'top': -4160, 'center': -4108, 'bottom': -4107}.items():
            self.setUp()
            self.assertEqual(num, self.apply(valign=name).writes['VerticalAlignment'],
                             f"valign: {name}")

    def test_an_unknown_alignment_is_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(align='middle')
        self.assertIn('align', str(ctx.exception))
        self.assertIn('center', str(ctx.exception))

    # --- number_format --------------------------------------------------

    def test_a_format_code_is_passed_straight_through(self):
        self.apply(number_format='#,##0')
        self.assertEqual('#,##0', self.ole.writes['NumberFormat'])
        self.assertEqual([], self.ole.application.international_calls)

    # 'General' is the one code that cannot be written: measured, it fails
    # on a localized Excel, whose own spelling differs.
    # Application.International(26) gives the right spelling on any locale.
    def test_general_goes_through_international_26(self):
        self.apply(number_format='general')
        self.assertEqual([26], self.ole.application.international_calls)
        self.assertEqual(LOCAL_GENERAL, self.ole.writes['NumberFormat'])

    def test_the_string_general_takes_the_same_route_whatever_its_case(self):
        for spelling in ['General', 'general', 'GENERAL']:
            self.setUp()
            self.apply(number_format=spelling)
            self.assertEqual([26], self.ole.application.international_calls, spelling)
            self.assertNotEqual(spelling, self.ole.writes['NumberFormat'],
                                f"{spelling!r} must not be sent as-is -- it fails against "
                                'real Excel')

    def test_text_names_the_at_sign_format(self):
        # Ruby spells it `:text`; here the string 'text' (any case) is the
        # same shorthand for the '@' format code, the way 'general' is a
        # shorthand -- and the code itself still passes straight through.
        for spelling in ('text', 'Text', 'TEXT', '@'):
            self.setUp()
            self.apply(number_format=spelling)
            self.assertEqual('@', self.ole.writes['NumberFormat'])
            self.assertEqual([], self.ole.application.international_calls)

    def test_a_non_string_number_format_is_refused(self):
        with self.assertRaises(TypeError) as ctx:
            self.apply(number_format=3)
        self.assertIn('number_format', str(ctx.exception))

    # --- values are checked too, not just key names -----------------------

    # None means "I do not have a value for this", which is the same thing
    # as not passing the key. It must NOT reach COM: measured, a COM boolean
    # assigned nil reads back as false, which would be the opposite of what
    # the caller meant.
    def test_none_means_the_key_was_not_specified(self):
        self.apply(bold=None, italic=True, wrap=None)
        self.assertNotIn('Bold', self.ole.font.writes,
                         'bold=None must be treated as "not specified"')
        self.assertNotIn('WrapText', self.ole.writes)
        self.assertIs(True, self.ole.font.writes['Italic'],
                      'the keys that do have values still apply')

    def test_a_call_of_nothing_but_nones_touches_nothing(self):
        self.apply(bold=None, background=None, align=None)
        self.assertEqual(0, self.ole.font_fetches)
        self.assertEqual(0, self.ole.interior_fetches)
        self.assertEqual({}, self.ole.writes)

    # A bad value must leave the range exactly as it was, the same way a
    # misspelled key does -- a caller who sees an exception reads it as
    # "nothing happened".
    def test_a_bad_value_is_caught_before_anything_is_written(self):
        with self.assertRaises(ValueError):
            self.apply(bold=True, align='middle')
        self.assertEqual(0, self.ole.font_fetches,
                         'validation must finish before the first COM call, or the range '
                         'is left half-formatted')

    def test_a_non_boolean_is_refused(self):
        for bad in [0, 1, 'true', 'yes']:
            self.setUp()
            with self.assertRaises(TypeError, msg=f"bold={bad!r} must be refused"):
                self.apply(bold=bad)

    def test_a_bad_size_is_refused(self):
        for bad, expected in ((0, ValueError), (-1, ValueError), (410, ValueError),
                              ('big', TypeError), (math.inf, ValueError)):
            self.setUp()
            with self.assertRaises(expected, msg=f"size={bad!r} must be refused"):
                self.apply(size=bad)

    def test_a_valid_size_passes_through(self):
        self.assertEqual(14.5, self.apply(size=14.5).font.writes['Size'])

    # --- unknown keys ----------------------------------------------------

    def test_a_misspelled_key_is_refused_rather_than_ignored(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(**{'blod': True})
        self.assertIn('blod', str(ctx.exception))
        self.assertIn('bold', str(ctx.exception),
                      'the message must list the keys that do exist')

    def test_a_misspelled_key_is_caught_before_anything_is_written(self):
        with self.assertRaises(ValueError):
            self.apply(**{'bold': True, 'blod': True})
        self.assertEqual(0, self.ole.font_fetches,
                         'validation must happen before any COM call, so a typo leaves '
                         'the sheet untouched')

    # --- round trips ------------------------------------------------------

    def test_font_is_fetched_once_however_many_font_keys(self):
        self.apply(bold=True, italic=True, underline=True, size=12, color='#000000')
        self.assertEqual(1, self.ole.font_fetches,
                         'each ole.Font() is its own round trip -- fetch it once and reuse it')

    def test_interior_is_not_fetched_when_no_interior_key_is_given(self):
        self.apply(bold=True)
        self.assertEqual(0, self.ole.interior_fetches)

    # --- absent means untouched, for every key ------------------------------

    # test_an_absent_key_touches_nothing above proves this for `bold` alone.
    # Every other key could be protected only by coincidence: a missing
    # `if key in opts` guard could write a default straight through (e.g.
    # WrapText = False on every call) and nothing but this matrix would
    # notice, since a call with other keys present -- the shape a real
    # caller uses -- would not trip the all-empty test either.
    #
    # Each case supplies a filler option so the COM object the key lives on
    # is actually exercised, the way a real call would.
    ABSENT_KEY_CASES = {
        'bold': ({'italic': True}, 'font', ['Bold']),
        'italic': ({'bold': True}, 'font', ['Italic']),
        'underline': ({'bold': True}, 'font', ['Underline']),
        'size': ({'bold': True}, 'font', ['Size']),
        'color': ({'bold': True}, 'font', ['Color', 'ColorIndex']),
        'background': ({'bold': True}, 'interior', ['Color', 'ColorIndex']),
        'align': ({'bold': True}, 'range', ['HorizontalAlignment']),
        'valign': ({'bold': True}, 'range', ['VerticalAlignment']),
        'wrap': ({'bold': True}, 'range', ['WrapText']),
        'number_format': ({'bold': True}, 'range', ['NumberFormat']),
    }

    def _writes_of(self, where):
        if where == 'font':
            return self.ole.font.writes
        if where == 'interior':
            return self.ole.interior.writes
        return self.ole.writes

    def test_an_absent_key_touches_nothing_for_every_key(self):
        for key, (filler, where, com_keys) in self.ABSENT_KEY_CASES.items():
            self.setUp()
            self.apply(**filler)
            for com_key in com_keys:
                self.assertNotIn(com_key, self._writes_of(where),
                                 f"{key!r} was omitted, so {com_key} must not have been "
                                 f"written (filler was {filler!r})")

    def test_the_font_keys_actually_fetch_font_via_their_filler(self):
        for key in ['italic', 'underline', 'size', 'color', 'bold']:
            self.setUp()
            self.apply(**self.ABSENT_KEY_CASES[key][0])
            self.assertEqual(1, self.ole.font_fetches,
                             f"the filler for {key!r} must exercise Font, or the matrix "
                             'test above proves nothing')

    # --- border -----------------------------------------------------------

    # 'all' resolves to every edge, so this takes the bulk path: an
    # assignment straight to Borders itself, which Excel applies to all six
    # edges -- including inside_v and inside_h -- rather than a per-edge
    # Item() loop.
    def test_border_all_sets_every_edge_including_the_inside_ones(self):
        self.apply(border='all')
        self.assertEqual(1, self.ole.borders.writes['LineStyle'])
        self.assertEqual(2, self.ole.borders.writes['Weight'],
                         'the simple form means a thin continuous line')
        self.assertEqual(0, self.ole.borders.fetches, 'the bulk path never calls Item()')
        self.assertEqual({}, self.ole.borders.items)

    def test_border_outline_leaves_the_inside_edges_alone(self):
        self.apply(border='outline')
        self.assertEqual([7, 8, 9, 10], sorted(self.ole.borders.items))
        self.assertEqual(4, self.ole.borders.fetches,
                         "'outline' must never take the bulk path -- that would draw the "
                         'inside edges too')

    def test_a_single_edge(self):
        self.apply(border='bottom')
        self.assertEqual([9], list(self.ole.borders.items))

    def test_a_list_of_edges(self):
        self.apply(border=['top', 'bottom'])
        self.assertEqual([8, 9], sorted(self.ole.borders.items))

    # False also resolves to every edge, so it takes the bulk path too.
    def test_border_false_clears_every_edge(self):
        self.apply(border=False)
        self.assertEqual(-4142, self.ole.borders.writes['LineStyle'])
        self.assertNotIn('Weight', self.ole.borders.writes,
                         'clearing must not set a weight -- there is no line to weigh')

    def test_the_dict_form_carries_style_and_colour(self):
        self.apply(border={'edges': 'outline', 'style': 'medium', 'color': '#999999'})
        self.assertEqual([7, 8, 9, 10], sorted(self.ole.borders.items))
        b = self.ole.borders.items[7]
        self.assertEqual(1, b.writes['LineStyle'])
        self.assertEqual(-4138, b.writes['Weight'])
        self.assertEqual(Color.parse('#999999'), b.writes['Color'])

    def test_every_border_style(self):
        for name, (line, weight) in {'hairline': (1, 1), 'thin': (1, 2),
                                     'medium': (1, -4138), 'thick': (1, 4),
                                     'dash': (-4115, 2), 'dot': (-4118, 2)}.items():
            self.setUp()
            self.apply(border={'edges': 'bottom', 'style': name})
            b = self.ole.borders.items[9]
            self.assertEqual(line, b.writes['LineStyle'], f"style {name} line")
            self.assertEqual(weight, b.writes['Weight'], f"style {name} weight")

    def test_the_dict_form_defaults_to_thin(self):
        self.apply(border={'edges': 'bottom'})
        self.assertEqual(2, self.ole.borders.items[9].writes['Weight'])

    # 'none' has no line to colour: the colour is still validated (a bad
    # one raises), but there is deliberately nothing left to write it to.
    def test_none_style_with_a_colour_validates_but_never_writes_the_colour(self):
        self.apply(border={'edges': 'all', 'style': 'none', 'color': '#FF0000'})
        self.assertEqual(-4142, self.ole.borders.writes['LineStyle'])
        self.assertNotIn('Weight', self.ole.borders.writes)
        self.assertNotIn('Color', self.ole.borders.writes,
                         "style 'none' has no line to colour, so Color must never be written")

    # The error for a bad edge says "a list of those", and "those" includes
    # 'all' and 'outline' -- so both must actually work inside a list.
    def test_all_expands_inside_a_list(self):
        self.apply(border=['all'])
        self.assertEqual(1, self.ole.borders.writes['LineStyle'])
        self.assertEqual(0, self.ole.borders.fetches,
                         "a list containing only 'all' still takes the bulk path")

    def test_outline_expands_inside_a_list(self):
        self.apply(border=['outline'])
        self.assertEqual([7, 8, 9, 10], sorted(self.ole.borders.items))

    def test_outline_can_be_combined_with_another_edge_in_a_list(self):
        self.apply(border=['outline', 'inside_h'])
        self.assertEqual([7, 8, 9, 10, 12], sorted(self.ole.borders.items))

    # An empty edge list writes nothing, so it must not even fetch Borders.
    def test_an_empty_edge_list_fetches_nothing(self):
        self.apply(border=[])
        self.assertEqual(0, self.ole.borders_fetches)

    def test_duplicate_edges_are_written_once(self):
        self.apply(border=['top', 'top'])
        self.assertEqual(1, self.ole.borders.fetches,
                         'a duplicated edge must be deduplicated before writing')

    def test_borders_is_fetched_once(self):
        self.apply(border='all')
        self.assertEqual(1, self.ole.borders_fetches,
                         'each ole.Borders() is a round trip -- fetch it once and Item() off it')

    def test_an_unknown_edge_is_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(border='diagonal')
        self.assertIn('border', str(ctx.exception))
        self.assertIn('outline', str(ctx.exception))

    def test_an_unknown_border_style_is_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(border={'edges': 'all', 'style': 'squiggly'})
        self.assertIn('style', str(ctx.exception))
        self.assertIn('thin', str(ctx.exception))

    def test_a_raw_colour_integer_in_a_border_is_refused(self):
        with self.assertRaises(ValueError):
            self.apply(border={'edges': 'all', 'color': 255})

    def test_an_unknown_key_inside_the_border_dict_is_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(border={'edges': 'all', 'widht': 2})
        self.assertIn('widht', str(ctx.exception))

    # None means "no value for this" everywhere else in this module -- an
    # explicit style=None inside the border dict must fall back to 'thin'
    # rather than override the default with a value that then fails to
    # validate.
    def test_none_inside_the_border_dict_falls_back_to_the_default(self):
        self.apply(border={'edges': 'all', 'style': None})
        self.assertEqual(2, self.ole.borders.writes['Weight'],
                         "style=None must fall back to 'thin'")

    def test_border_none_means_not_specified(self):
        self.apply(border=None, bold=True)
        self.assertEqual(0, self.ole.borders_fetches,
                         'border=None must be treated as "not specified"')
        self.assertIs(True, self.ole.font.writes['Bold'])

    # Same discipline as every other key: a bad value leaves the range alone.
    def test_a_bad_border_is_caught_before_anything_is_written(self):
        with self.assertRaises(ValueError):
            self.apply(bold=True, border='diagonal')
        self.assertEqual(0, self.ole.borders_fetches)
        self.assertEqual(0, self.ole.font_fetches,
                         'nothing may be written until the whole option set has been validated')

    def test_a_border_of_an_unsupported_type_is_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(border=3)
        self.assertIn('border', str(ctx.exception))

    # --- the key-name check runs before the None drop ----------------------

    # A misspelled key whose value happens to be None would be silently
    # absorbed if apply dropped Nones before checking key names. Key names
    # must be checked against the options as given.
    def test_a_misspelled_key_with_a_none_value_is_still_refused(self):
        with self.assertRaises(ValueError) as ctx:
            self.apply(**{'blod': None})
        self.assertIn('blod', str(ctx.exception))

    def test_a_correct_key_with_a_none_value_is_still_a_no_op(self):
        self.apply(bold=None)
        self.assertEqual({}, self.ole.font.writes)
        self.assertEqual(0, self.ole.font_fetches)


if __name__ == '__main__':
    unittest.main()
