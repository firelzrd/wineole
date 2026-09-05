require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/address'

class AddressTest < Minitest::Test
  A = WineOLE::MSOffice::Address

  def parse(s, version = 11.0)
    A.parse(s, version)
  end

  def test_parses_a_bare_range
    a = parse('A1:B2')
    assert_nil a.workbook
    assert_nil a.worksheet
    assert_equal 'A1:B2', a.range
    assert a.range?
  end

  def test_parses_a_worksheet_and_range
    a = parse('Sheet1!A1')
    assert_equal 'Sheet1', a.worksheet
    assert_equal 'A1', a.range
  end

  def test_parses_a_workbook_worksheet_and_range
    a = parse('[Book1]Sheet1!A1:B2')
    assert_equal 'Book1', a.workbook
    assert_equal 'Sheet1', a.worksheet
    assert_equal 'A1:B2', a.range
  end

  def test_parses_the_new_workbook_marker
    a = parse('[:new]')
    assert_equal ':new', a.workbook
    assert_nil a.range
    refute a.range?
  end

  def test_parses_the_sheet_markers
    assert_equal ':new',   parse(':new!').worksheet
    assert_equal ':first', parse(':first!A1').worksheet
    assert_equal ':last',  parse(':last!A1').worksheet
  end

  def test_an_empty_workbook_or_worksheet_means_the_active_one
    assert_equal '', parse('[]').workbook
    assert_equal '', parse('!A1').worksheet
  end

  def test_a_sheet_without_a_range_is_not_assignable
    a = parse('Sheet1!')
    assert_equal 'Sheet1', a.worksheet
    assert_nil a.range
    refute a.range?, 'a sheet address carries no range, so it must not accept an assignment'
  end

  def test_rejects_a_row_beyond_the_excel_11_limit
    assert_nil parse('A65537'), 'row 65537 does not exist in Excel 11'
    refute_nil parse('A65536')
  end

  def test_excel_12_allows_the_larger_grid
    assert_nil    parse('A65537', 11.0)
    refute_nil    parse('A65537', 12.0)
    assert_equal 'XFD1048576', parse('XFD1048576', 12.0).range
    assert_nil    parse('XFE1', 12.0), 'XFE is past the last column'
  end

  # Checking a handful of boundaries would not have caught what was actually
  # wrong with these patterns: row 65530 alone was missing from the Excel 11
  # grid, and 11,111 scattered rows from the Excel 12 one. Both grids are
  # small enough to check in full, and it takes well under a second.
  def test_every_row_in_the_grid_parses
    {11.0 => 65_536, 12.0 => 1_048_576}.each do |version, last_row|
      missing = (1..last_row).reject { |n| parse("A#{n}", version) }
      assert_empty missing.first(10),
        "#{missing.length} row(s) of the Excel #{version.to_i} grid do not parse"

      past_the_end = (last_row + 1..last_row + 20).select { |n| parse("A#{n}", version) }
      assert_empty past_the_end,
        "rows past the end of the Excel #{version.to_i} grid must not parse"
    end
  end

  def test_every_column_in_the_grid_parses
    columns = ->(n) {
      out = []
      ('A'..'Z').each { |a| out << a }
      ('A'..'Z').each { |a| ('A'..'Z').each { |b| out << a + b } }
      ('A'..'Z').each { |a| ('A'..'Z').each { |b| ('A'..'Z').each { |c| out << a + b + c } } }
      out.first(n)
    }

    {[11.0, 'IV'] => 256, [12.0, 'XFD'] => 16_384}.each do |(version, last), count|
      cols = columns.call(count)
      assert_equal last, cols.last, 'the column list must end at the grid limit'
      missing = cols.reject { |c| parse("#{c}1", version) }
      assert_empty missing.first(10),
        "#{missing.length} column(s) of the Excel #{version.to_i} grid do not parse"
    end
  end

  def test_returns_nil_for_something_that_is_not_an_address
    assert_nil parse('this is not an address'),
      'the caller treats nil as "not an address" and falls back to a raw sheet name'
  end
end
