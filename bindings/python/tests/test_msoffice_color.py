import math
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.msoffice.color import Color


class MSOfficeColorTest(unittest.TestCase):
    # Excel's Color is BGR. Red and blue are the only pair that can catch a
    # reversed implementation -- green is symmetric and would pass either
    # way, which is why it is not the example here.
    def test_red_and_blue_are_not_swapped(self):
        self.assertEqual(255, Color.parse('#FF0000'), 'red must be 255; 0xFF0000 would be blue')
        self.assertEqual(16711680, Color.parse('#0000FF'))

    def test_green_and_the_greys(self):
        self.assertEqual(65280, Color.parse('#00FF00'))
        self.assertEqual(0, Color.parse('#000000'))
        self.assertEqual(16777215, Color.parse('#FFFFFF'))
        self.assertEqual(8421504, Color.parse('#808080'))

    def test_the_hash_is_optional(self):
        self.assertEqual(Color.parse('#FF0000'), Color.parse('FF0000'))

    def test_three_digit_shorthand_doubles_each_digit(self):
        self.assertEqual(Color.parse('#EEEEEE'), Color.parse('#EEE'))
        self.assertEqual(Color.parse('#FF0000'), Color.parse('#F00'))

    def test_an_rgb_sequence(self):
        self.assertEqual(Color.parse('#FF8000'), Color.parse([255, 128, 0]))
        self.assertEqual(Color.parse('#FF8000'), Color.parse((255, 128, 0)),
                         'a tuple is the same three numbers as a list')

    def test_to_hex_is_the_inverse(self):
        for hex_value in ['#FF0000', '#0000FF', '#00FF00', '#123456', '#000000', '#FFFFFF']:
            self.assertEqual(hex_value, Color.to_hex(Color.parse(hex_value)),
                             f"{hex_value} must survive the round trip")

    # Excel hands back a float, not an int (measured: 255.0).
    def test_to_hex_accepts_the_float_excel_returns(self):
        self.assertEqual('#FF0000', Color.to_hex(255.0))

    # int(x, 16) would silently return 0 for these rather than complaining.
    def test_garbage_is_refused_rather_than_read_as_zero(self):
        for bad in ['#GGGGGG', '#12345', '#1234567', 'red', '', '#']:
            with self.assertRaises(ValueError, msg=f"{bad!r} must be refused") as ctx:
                Color.parse(bad)
            self.assertIn('#RRGGBB', str(ctx.exception))

    def test_a_bad_sequence_is_refused(self):
        for bad in [[255, 0], [255, 0, 0, 0], [255, 0, 256], [255, 0, -1], ['ff', 0, 0]]:
            with self.assertRaises(ValueError, msg=f"{bad!r} must be refused"):
                Color.parse(bad)

    def test_an_unsupported_type_is_refused(self):
        for bad in [None, 255, 1.5, {}, True]:
            with self.assertRaises(TypeError, msg=f"{bad!r} must be refused"):
                Color.parse(bad)

    def test_to_hex_refuses_a_non_number(self):
        with self.assertRaises(TypeError):
            Color.to_hex('#FF0000')

    # Font.ColorIndex returns -4105 for automatic, Interior.ColorIndex
    # returns -4142 for no fill. These are not colours and must be
    # rejected, not masked to a confident, meaningless answer.
    def test_to_hex_refuses_colorindex_automatic(self):
        with self.assertRaises(ValueError) as ctx:
            Color.to_hex(-4105)
        self.assertIn('ColorIndex', str(ctx.exception))

    def test_to_hex_refuses_colorindex_no_fill(self):
        with self.assertRaises(ValueError):
            Color.to_hex(-4142)

    def test_to_hex_refuses_overflow(self):
        with self.assertRaises(ValueError):
            Color.to_hex(0x1000000)

    # Out-of-range values must be caught before int() runs: int(inf) raises
    # OverflowError, which is neither of this module's two exceptions.
    def test_to_hex_refuses_infinity_as_a_value_error(self):
        with self.assertRaises(ValueError):
            Color.to_hex(math.inf)

    # complex is a number in Ruby's sense but not one this converter can
    # order against 0..0xFFFFFF, so it is refused as a wrong TYPE.
    def test_to_hex_refuses_complex_as_a_type_error(self):
        with self.assertRaises(TypeError):
            Color.to_hex(complex(1, 2))

    def test_to_hex_refuses_a_bool(self):
        # True is an int in Python. It is not a colour, and accepting it
        # would answer '#000001' with total confidence.
        with self.assertRaises(TypeError):
            Color.to_hex(True)

    def test_to_hex_accepts_boundary_zero(self):
        self.assertEqual('#000000', Color.to_hex(0))

    def test_to_hex_accepts_boundary_ffffff(self):
        self.assertEqual('#FFFFFF', Color.to_hex(0xFFFFFF))

    def test_a_sequence_with_float_elements_is_refused(self):
        with self.assertRaises(ValueError):
            Color.parse([255.0, 0, 0])

    def test_a_sequence_with_string_elements_is_refused(self):
        with self.assertRaises(ValueError):
            Color.parse(['255', '0', '0'])


if __name__ == '__main__':
    unittest.main()
