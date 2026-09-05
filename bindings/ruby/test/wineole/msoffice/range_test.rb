require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/range'

# Stands in for the COM Range. Records what was assigned to Value so the
# tests can assert on the exact shape that would have crossed the wire.
class FakeComRange
  attr_reader :written
  attr_accessor :value
  attr_reader :wrap_text

  def initialize(rows:, cols:, value: nil)
    @rows = rows
    @cols = cols
    @value = value
  end

  def Rows;    Struct.new(:Count).new(@rows); end
  def Columns; Struct.new(:Count).new(@cols); end
  def Value;   @value; end
  def Value=(v); @written = v; end
  def Interior; :the_com_interior; end

  def WrapText=(v); @wrap_text = v; end
end

class MSOfficeRangeTest < Minitest::Test
  def range(rows:, cols:, value: nil)
    WineOLE::MSOffice::Range.new(FakeComRange.new(rows: rows, cols: cols, value: value))
  end

  # --- to_a -------------------------------------------------------------
  def test_to_a_wraps_a_scalar_from_a_single_cell
    # Excel returns a bare scalar for a one-cell range, whatever the address.
    assert_equal [['a1']], range(rows: 1, cols: 1, value: 'a1').to_a
  end

  def test_to_a_passes_a_two_dimensional_value_through
    v = [[1, 2], [3, 4]]
    assert_equal v, range(rows: 2, cols: 2, value: v).to_a
  end

  def test_to_a_wraps_an_empty_single_cell
    assert_equal [[nil]], range(rows: 1, cols: 1, value: nil).to_a
  end

  # --- write: accepted --------------------------------------------------
  def test_write_broadcasts_a_scalar
    r = range(rows: 3, cols: 3)
    r.write(7)
    assert_equal 7, r.ole.written, 'Excel broadcasts a scalar itself; pass it through'
  end

  def test_write_orients_a_flat_array_down_a_column
    # Excel's own behaviour here replicates the first element down the column;
    # this is the trap `write` exists to close.
    r = range(rows: 3, cols: 1)
    r.write([1, 2, 3])
    assert_equal [[1], [2], [3]], r.ole.written
  end

  def test_write_orients_a_flat_array_across_a_row
    r = range(rows: 1, cols: 3)
    r.write([1, 2, 3])
    assert_equal [[1, 2, 3]], r.ole.written
  end

  def test_write_passes_a_matching_two_dimensional_array_through
    r = range(rows: 2, cols: 3)
    r.write([[1, 2, 3], [4, 5, 6]])
    assert_equal [[1, 2, 3], [4, 5, 6]], r.ole.written
  end

  # --- write: rejected --------------------------------------------------
  def test_write_rejects_a_flat_array_of_the_wrong_length
    err = assert_raises(ArgumentError) { range(rows: 3, cols: 1).write([1, 2]) }
    assert_match(/3x1/, err.message, 'the message must state the range size')
    assert_match(/2/, err.message, 'and what it was given')
  end

  def test_write_rejects_a_flat_array_when_the_range_is_not_a_line
    err = assert_raises(ArgumentError) { range(rows: 3, cols: 3).write([1, 2, 3]) }
    assert_match(/3x3/, err.message)
  end

  def test_write_rejects_a_two_dimensional_array_of_the_wrong_size
    err = assert_raises(ArgumentError) { range(rows: 3, cols: 3).write([[1, 2, 3, 4, 5]]) }
    assert_match(/3x3/, err.message)
    assert_match(/1x5/, err.message)
  end

  def test_write_rejects_a_ragged_array
    err = assert_raises(ArgumentError) { range(rows: 2, cols: 2).write([[1, 2], [3]]) }
    assert_match(/ragged/, err.message)
  end

  # --- fill -------------------------------------------------------------
  def test_fill_replicates_a_scalar_over_the_whole_range
    r = range(rows: 3, cols: 3)
    r.fill(7)
    assert_equal [[7, 7, 7], [7, 7, 7], [7, 7, 7]], r.ole.written
  end

  def test_fill_treats_a_flat_array_as_a_row_and_replicates_it_down
    r = range(rows: 3, cols: 3)
    r.fill([1, 2])
    assert_equal [[1, 2, nil], [1, 2, nil], [1, 2, nil]], r.ole.written
  end

  def test_fill_truncates_and_pads_a_two_dimensional_array
    r = range(rows: 3, cols: 3)
    r.fill([[1, 2, 3, 4, 5]])
    assert_equal [[1, 2, 3], [nil, nil, nil], [nil, nil, nil]], r.ole.written
  end

  # --- passthrough ------------------------------------------------------
  def test_unknown_methods_go_to_com
    assert_equal :the_com_interior, range(rows: 1, cols: 1).Interior
  end

  def test_value_is_not_intercepted
    # `Value` must stay exactly Excel's -- scalar for one cell -- so that
    # Excel documentation and VBA knowledge keep applying.
    assert_equal 'a1', range(rows: 1, cols: 1, value: 'a1').Value
  end

  def test_to_ary_is_not_defined
    # to_a is an explicit conversion and is ours; to_ary is the implicit one
    # `puts` and destructuring reach for, and Proxy deliberately leaves it
    # undefined. Defining it here would undo that.
    refute range(rows: 1, cols: 1).respond_to?(:to_ary)
  end

  # --- format delegates and returns self --------------------------------

  def test_format_returns_the_range_itself
    r = range(rows: 1, cols: 1)
    assert_same r, r.format(wrap: true)
  end

  def test_format_reaches_the_com_object
    fake = FakeComRange.new(rows: 1, cols: 1)
    WineOLE::MSOffice::Range.new(fake).format(wrap: true)
    assert_equal true, fake.wrap_text
  end

  def test_format_refuses_an_unknown_key
    err = assert_raises(ArgumentError) { range(rows: 1, cols: 1).format(blod: true) }
    assert_match(/blod/, err.message)
  end
end
