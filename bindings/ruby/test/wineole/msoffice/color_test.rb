require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/color'

class MSOfficeColorTest < Minitest::Test
  C = WineOLE::MSOffice::Color

  # Excel's Color is BGR. Red and blue are the only pair that can catch a
  # reversed implementation -- green is symmetric and would pass either way,
  # which is why it is not the example here.
  def test_red_and_blue_are_not_swapped
    assert_equal 255, C['#FF0000'], 'red must be 255; 0xFF0000 would be blue'
    assert_equal 16_711_680, C['#0000FF']
  end

  def test_green_and_the_greys
    assert_equal 65_280, C['#00FF00']
    assert_equal 0, C['#000000']
    assert_equal 16_777_215, C['#FFFFFF']
    assert_equal 8_421_504, C['#808080']
  end

  def test_the_hash_is_optional
    assert_equal C['#FF0000'], C['FF0000']
  end

  def test_three_digit_shorthand_doubles_each_digit
    assert_equal C['#EEEEEE'], C['#EEE']
    assert_equal C['#FF0000'], C['#F00']
  end

  def test_an_rgb_array
    assert_equal C['#FF8000'], C[[255, 128, 0]]
  end

  def test_to_hex_is_the_inverse
    ['#FF0000', '#0000FF', '#00FF00', '#123456', '#000000', '#FFFFFF'].each do |hex|
      assert_equal hex, C.to_hex(C[hex]), "#{hex} must survive the round trip"
    end
  end

  # Excel hands back a Float, not an Integer (measured: 255.0).
  def test_to_hex_accepts_the_float_excel_returns
    assert_equal '#FF0000', C.to_hex(255.0)
  end

  # to_i(16) would silently return 0 for these rather than complaining.
  def test_garbage_is_refused_rather_than_read_as_zero
    ['#GGGGGG', '#12345', '#1234567', 'red', '', '#'].each do |bad|
      err = assert_raises(ArgumentError, "#{bad.inspect} must be refused") { C[bad] }
      assert_match(/#RRGGBB/, err.message)
    end
  end

  def test_a_bad_array_is_refused
    [[255, 0], [255, 0, 0, 0], [255, 0, 256], [255, 0, -1], ['ff', 0, 0]].each do |bad|
      assert_raises(ArgumentError, "#{bad.inspect} must be refused") { C[bad] }
    end
  end

  def test_an_unsupported_type_is_refused
    [nil, :red, 255, 1.5, {}].each do |bad|
      assert_raises(ArgumentError, "#{bad.inspect} must be refused") { C[bad] }
    end
  end

  def test_to_hex_refuses_a_non_number
    assert_raises(ArgumentError) { C.to_hex('#FF0000') }
  end

  # Font.ColorIndex returns -4105 for automatic, Interior.ColorIndex returns
  # -4142 for no fill. These are not colours and must be rejected, not masked
  # to a confident, meaningless answer.
  def test_to_hex_refuses_colorindex_automatic
    err = assert_raises(ArgumentError) { C.to_hex(-4105) }
    assert_match(/ColorIndex/, err.message)
  end

  def test_to_hex_refuses_colorindex_no_fill
    assert_raises(ArgumentError) { C.to_hex(-4142) }
  end

  def test_to_hex_refuses_overflow
    assert_raises(ArgumentError) { C.to_hex(0x1000000) }
  end

  # Out-of-range values must be caught before #to_i runs -- Float::INFINITY
  # raises FloatDomainError from #to_i, and Complex raises RangeError, and
  # neither is this module's ArgumentError.
  def test_to_hex_refuses_infinity_as_argument_error
    assert_raises(ArgumentError) { C.to_hex(Float::INFINITY) }
  end

  def test_to_hex_refuses_complex_as_argument_error
    assert_raises(ArgumentError) { C.to_hex(Complex(1, 2)) }
  end

  # to_hex must accept the boundaries 0 and 0xFFFFFF, and they must convert
  # correctly: 0 -> '#000000' (black), 0xFFFFFF -> '#FFFFFF' (white).
  def test_to_hex_accepts_boundary_zero
    assert_equal '#000000', C.to_hex(0)
  end

  def test_to_hex_accepts_boundary_ffffff
    assert_equal '#FFFFFF', C.to_hex(0xFFFFFF)
  end

  # from_array already checks is_a?(::Integer), but no test documents it,
  # and a future simplification to is_a?(::Numeric) would pass the suite.
  # This documents that float elements are refused.
  def test_array_with_float_elements_is_refused
    assert_raises(ArgumentError) { C[[255.0, 0, 0]] }
  end

  def test_array_with_string_elements_is_refused
    assert_raises(ArgumentError) { C[['255', '0', '0']] }
  end
end
