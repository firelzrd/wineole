require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/format'

# Records every COM assignment so a test can assert not only what was set
# but what was NOT touched -- which is the whole point of the three-state
# rule. Font and Interior are separate objects, as they are in COM, and
# each counts how many times it was fetched so the round-trip saving can
# be asserted rather than assumed.
class FakeFont
  attr_reader :writes
  def initialize; @writes = {}; end
  def Bold=(v);       @writes[:Bold] = v; end
  def Italic=(v);     @writes[:Italic] = v; end
  def Underline=(v);  @writes[:Underline] = v; end
  def Size=(v);       @writes[:Size] = v; end
  def Color=(v);      @writes[:Color] = v; end
  def ColorIndex=(v); @writes[:ColorIndex] = v; end
end

class FakeInterior
  attr_reader :writes
  def initialize; @writes = {}; end
  def Color=(v);      @writes[:Color] = v; end
  def ColorIndex=(v); @writes[:ColorIndex] = v; end
end

# Whatever this Excel calls the General format. Deliberately not the value
# this project's own host returns: a real spelling here would bake one
# machine's locale into the suite and would let an implementation that
# hardcodes that spelling pass.
LOCAL_GENERAL = 'LOCAL-GENERAL-FORMAT-NAME'

class FakeApplication
  attr_reader :international_calls
  def initialize; @international_calls = []; end
  def International(index)
    @international_calls << index
    LOCAL_GENERAL
  end
end

class FakeBorder
  attr_reader :writes
  def initialize; @writes = {}; end
  def LineStyle=(v); @writes[:LineStyle] = v; end
  def Weight=(v);    @writes[:Weight] = v; end
  def Color=(v);     @writes[:Color] = v; end
end

class FakeBorders
  attr_reader :items, :fetches, :writes
  def initialize; @items = {}; @fetches = 0; @writes = {}; end
  def Item(index)
    @fetches += 1
    @items[index] ||= FakeBorder.new
  end
  def LineStyle=(v); @writes[:LineStyle] = v; end
  def Weight=(v);    @writes[:Weight] = v; end
  def Color=(v);     @writes[:Color] = v; end
end

class FakeComRangeForFormat
  attr_reader :writes, :font, :interior, :application, :font_fetches, :interior_fetches, :borders

  def initialize
    @writes = {}
    @font = FakeFont.new
    @interior = FakeInterior.new
    @application = FakeApplication.new
    @font_fetches = 0
    @interior_fetches = 0
    @borders = FakeBorders.new
  end

  def Font;        @font_fetches += 1; @font; end
  def Interior;    @interior_fetches += 1; @interior; end
  def Application; @application; end
  def Borders; @borders_fetches = (@borders_fetches || 0) + 1; @borders; end
  def borders_fetches; @borders_fetches || 0; end

  def HorizontalAlignment=(v); @writes[:HorizontalAlignment] = v; end
  def VerticalAlignment=(v);   @writes[:VerticalAlignment] = v; end
  def WrapText=(v);            @writes[:WrapText] = v; end
  def NumberFormat=(v);        @writes[:NumberFormat] = v; end
end

class MSOfficeFormatTest < Minitest::Test
  F = WineOLE::MSOffice::Format

  def setup
    @ole = FakeComRangeForFormat.new
  end

  def apply(**opts)
    F.apply(@ole, opts)
    @ole
  end

  # --- the three-state rule ------------------------------------------

  def test_an_absent_key_touches_nothing
    apply(italic: true)
    refute @ole.font.writes.key?(:Bold),
      'a key that was not passed must not be written at all -- absent means "leave it alone"'
  end

  def test_false_is_written_rather_than_skipped
    apply(bold: false)
    assert_equal false, @ole.font.writes[:Bold],
      'false means "explicitly turn it off", which is not the same as leaving it alone'
  end

  def test_nothing_at_all_touches_nothing
    apply
    assert_empty @ole.writes
    assert_equal 0, @ole.font_fetches
    assert_equal 0, @ole.interior_fetches
  end

  # --- font ----------------------------------------------------------

  def test_bold_and_italic_and_size
    apply(bold: true, italic: true, size: 14)
    assert_equal true, @ole.font.writes[:Bold]
    assert_equal true, @ole.font.writes[:Italic]
    assert_equal 14, @ole.font.writes[:Size]
  end

  # Font.Underline is not a boolean: measured, its default is -4142 and
  # single underline is 2. Sending `true` would mean something else.
  def test_underline_true_becomes_single_not_true
    apply(underline: true)
    assert_equal 2, @ole.font.writes[:Underline]
  end

  def test_underline_false_becomes_none
    apply(underline: false)
    assert_equal(-4142, @ole.font.writes[:Underline])
  end

  def test_underline_symbols
    assert_equal 2, apply(underline: :single).font.writes[:Underline]
    setup
    assert_equal(-4119, apply(underline: :double).font.writes[:Underline])
    setup
    assert_equal(-4142, apply(underline: :none).font.writes[:Underline])
  end

  def test_an_unknown_underline_symbol_is_refused
    err = assert_raises(ArgumentError) { apply(underline: :wavy) }
    assert_match(/underline/, err.message)
    assert_match(/single/, err.message)
  end

  # --- colour --------------------------------------------------------

  def test_colour_goes_through_the_bgr_conversion
    apply(color: '#FF0000', background: '#0000FF')
    assert_equal 255, @ole.font.writes[:Color], 'red must be 255, not 0xFF0000'
    assert_equal 16_711_680, @ole.interior.writes[:Color]
  end

  # Color cannot express "no fill": a cleared cell and a white-painted cell
  # both read back 16777215. Clearing has to go through ColorIndex.
  def test_background_false_clears_rather_than_painting_white
    apply(background: false)
    assert_equal(-4142, @ole.interior.writes[:ColorIndex])
    refute @ole.interior.writes.key?(:Color),
      'clearing must not write a Color at all -- painting white leaves the cell filled'
  end

  def test_color_false_sets_automatic
    apply(color: false)
    assert_equal(-4105, @ole.font.writes[:ColorIndex])
    refute @ole.font.writes.key?(:Color)
  end

  def test_a_raw_com_colour_integer_is_refused
    [255, 0xFF0000, 0].each do |raw|
      err = assert_raises(ArgumentError, "#{raw} must be refused") { F.apply(FakeComRangeForFormat.new, {color: raw}) }
      assert_match(/Color/, err.message, 'the message must point at MSOffice::Color')
    end
  end

  # --- alignment, wrap ------------------------------------------------

  def test_alignment_symbols_become_the_measured_numbers
    apply(align: :center, valign: :top, wrap: true)
    assert_equal(-4108, @ole.writes[:HorizontalAlignment])
    assert_equal(-4160, @ole.writes[:VerticalAlignment])
    assert_equal true, @ole.writes[:WrapText]
  end

  def test_every_alignment_symbol
    {general: 1, left: -4131, center: -4108, right: -4152, justify: -4130}.each do |sym, num|
      setup
      assert_equal num, apply(align: sym).writes[:HorizontalAlignment], "align: #{sym}"
    end
    {top: -4160, center: -4108, bottom: -4107}.each do |sym, num|
      setup
      assert_equal num, apply(valign: sym).writes[:VerticalAlignment], "valign: #{sym}"
    end
  end

  def test_an_unknown_alignment_is_refused
    err = assert_raises(ArgumentError) { apply(align: :middle) }
    assert_match(/align/, err.message)
    assert_match(/center/, err.message)
  end

  # --- number_format --------------------------------------------------

  def test_a_format_code_is_passed_straight_through
    apply(number_format: '#,##0')
    assert_equal '#,##0', @ole.writes[:NumberFormat]
    assert_empty @ole.application.international_calls
  end

  # 'General' is the one code that cannot be written: measured, it fails on
  # a localized Excel, whose own spelling differs. Application.International(26)
  # gives the right spelling on any locale.
  def test_general_goes_through_international_26
    apply(number_format: :general)
    assert_equal [26], @ole.application.international_calls
    assert_equal LOCAL_GENERAL, @ole.writes[:NumberFormat]
  end

  def test_the_string_general_takes_the_same_route_whatever_its_case
    ['General', 'general', 'GENERAL'].each do |spelling|
      setup
      apply(number_format: spelling)
      assert_equal [26], @ole.application.international_calls, spelling
      refute_equal spelling, @ole.writes[:NumberFormat],
        "#{spelling.inspect} must not be sent as-is -- it fails against real Excel"
    end
  end

  def test_text_becomes_the_at_sign
    apply(number_format: :text)
    assert_equal '@', @ole.writes[:NumberFormat]
  end

  def test_an_unknown_number_format_symbol_is_refused
    err = assert_raises(ArgumentError) { apply(number_format: :currency) }
    assert_match(/number_format/, err.message)
  end

  # --- values are checked too, not just key names -----------------------

  # nil means "I do not have a value for this", which is the same thing as
  # not passing the key. It must NOT reach COM: measured, a COM boolean
  # assigned nil reads back as false, which would be the opposite of what
  # the caller meant.
  def test_nil_means_the_key_was_not_specified
    apply(bold: nil, italic: true, wrap: nil)
    refute @ole.font.writes.key?(:Bold), 'bold: nil must be treated as "not specified"'
    refute @ole.writes.key?(:WrapText)
    assert_equal true, @ole.font.writes[:Italic], 'the keys that do have values still apply'
  end

  def test_a_hash_of_nothing_but_nils_touches_nothing
    apply(bold: nil, background: nil, align: nil)
    assert_equal 0, @ole.font_fetches
    assert_equal 0, @ole.interior_fetches
    assert_empty @ole.writes
  end

  # A bad value must leave the range exactly as it was, the same way a
  # misspelled key does -- a caller who sees an exception reads it as
  # "nothing happened".
  def test_a_bad_value_is_caught_before_anything_is_written
    assert_raises(ArgumentError) { apply(bold: true, align: :middle) }
    assert_equal 0, @ole.font_fetches,
      'validation must finish before the first COM call, or the range is left half-formatted'
  end

  def test_a_non_boolean_is_refused
    [0, 1, 'true', :yes].each do |bad|
      setup
      assert_raises(ArgumentError, "bold: #{bad.inspect} must be refused") { apply(bold: bad) }
    end
  end

  def test_a_bad_size_is_refused
    [0, -1, 410, 'big', Float::INFINITY].each do |bad|
      setup
      assert_raises(ArgumentError, "size: #{bad.inspect} must be refused") { apply(size: bad) }
    end
  end

  def test_a_valid_size_passes_through
    assert_equal 14.5, apply(size: 14.5).font.writes[:Size]
  end

  # --- unknown keys ----------------------------------------------------

  def test_a_misspelled_key_is_refused_rather_than_ignored
    err = assert_raises(ArgumentError) { apply(blod: true) }
    assert_match(/blod/, err.message)
    assert_match(/bold/, err.message, 'the message must list the keys that do exist')
  end

  def test_a_misspelled_key_is_caught_before_anything_is_written
    assert_raises(ArgumentError) { apply(bold: true, blod: true) }
    assert_equal 0, @ole.font_fetches,
      'validation must happen before any COM call, so a typo leaves the sheet untouched'
  end

  # --- round trips ------------------------------------------------------

  def test_font_is_fetched_once_however_many_font_keys
    apply(bold: true, italic: true, underline: true, size: 12, color: '#000000')
    assert_equal 1, @ole.font_fetches,
      'each `ole.Font` is its own round trip -- fetch it once and reuse it'
  end

  def test_interior_is_not_fetched_when_no_interior_key_is_given
    apply(bold: true)
    assert_equal 0, @ole.interior_fetches
  end

  # --- absent means untouched, for every key ------------------------------

  # test_an_absent_key_touches_nothing above proves this for :bold alone.
  # Every other key used to be protected only by coincidence: with nil
  # refused (before Change 1), a missing `if opts.key?(k)` guard left
  # opts[k] as nil, which some unrelated validator then rejected -- so the
  # test suite would still fail, just not for the right reason. Now that
  # nil is dropped before validation, that coincidence is gone: a missing
  # guard could write a default straight through (e.g. `WrapText = false`
  # on every call) and nothing but this matrix would notice, since a call
  # with other keys present -- the shape a real caller uses -- would not
  # trip the all-empty test either.
  #
  # Each case supplies a filler option so the COM object the key lives on
  # is actually exercised, the way a real call would, rather than skipped
  # entirely.
  ABSENT_KEY_CASES = {
    bold: {filler: {italic: true}, writes: ->(o) { o.font.writes }, keys: [:Bold]},
    italic: {filler: {bold: true}, writes: ->(o) { o.font.writes }, keys: [:Italic]},
    underline: {filler: {bold: true}, writes: ->(o) { o.font.writes }, keys: [:Underline]},
    size: {filler: {bold: true}, writes: ->(o) { o.font.writes }, keys: [:Size]},
    color: {filler: {bold: true}, writes: ->(o) { o.font.writes }, keys: %i[Color ColorIndex]},
    background: {filler: {bold: true}, writes: ->(o) { o.interior.writes }, keys: %i[Color ColorIndex]},
    align: {filler: {bold: true}, writes: ->(o) { o.writes }, keys: [:HorizontalAlignment]},
    valign: {filler: {bold: true}, writes: ->(o) { o.writes }, keys: [:VerticalAlignment]},
    wrap: {filler: {bold: true}, writes: ->(o) { o.writes }, keys: [:WrapText]},
    number_format: {filler: {bold: true}, writes: ->(o) { o.writes }, keys: [:NumberFormat]}
  }.freeze

  def test_an_absent_key_touches_nothing_for_every_key
    ABSENT_KEY_CASES.each do |key, spec|
      setup
      apply(**spec[:filler])
      spec[:keys].each do |com_key|
        refute spec[:writes].call(@ole).key?(com_key),
          "#{key.inspect} was omitted, so #{com_key} must not have been written " \
          "(filler was #{spec[:filler].inspect})"
      end
    end
  end

  def test_the_font_keys_actually_fetch_font_via_their_filler
    %i[italic underline size color].each do |key|
      setup
      apply(**ABSENT_KEY_CASES[key][:filler])
      assert_equal 1, @ole.font_fetches,
        "the filler for #{key.inspect} must exercise Font, or the matrix test above proves nothing"
    end
    setup
    apply(**ABSENT_KEY_CASES[:bold][:filler])
    assert_equal 1, @ole.font_fetches
  end

  # --- border -----------------------------------------------------------

  # :all resolves to every edge, so this takes the bulk path: an assignment
  # straight to Borders itself, which Excel applies to all six edges --
  # including inside_v and inside_h -- rather than a per-edge Item() loop.
  def test_border_all_sets_every_edge_including_the_inside_ones
    apply(border: :all)
    assert_equal 1, @ole.borders.writes[:LineStyle]
    assert_equal 2, @ole.borders.writes[:Weight], 'the simple form means a thin continuous line'
    assert_equal 0, @ole.borders.fetches, 'the bulk path never calls Item()'
    assert_empty @ole.borders.items
  end

  def test_border_outline_leaves_the_inside_edges_alone
    apply(border: :outline)
    assert_equal [7, 8, 9, 10].sort, @ole.borders.items.keys.sort
    assert_equal 4, @ole.borders.fetches,
      ':outline must never take the bulk path -- that would draw the inside edges too'
  end

  def test_a_single_edge
    apply(border: :bottom)
    assert_equal [9], @ole.borders.items.keys
  end

  def test_an_array_of_edges
    apply(border: [:top, :bottom])
    assert_equal [8, 9].sort, @ole.borders.items.keys.sort
  end

  # false also resolves to every edge (see normalize_border), so it takes
  # the bulk path too.
  def test_border_false_clears_every_edge
    apply(border: false)
    assert_equal(-4142, @ole.borders.writes[:LineStyle])
    refute @ole.borders.writes.key?(:Weight),
      'clearing must not set a weight -- there is no line to weigh'
  end

  def test_the_hash_form_carries_style_and_colour
    apply(border: {edges: :outline, style: :medium, color: '#999999'})
    assert_equal [7, 8, 9, 10].sort, @ole.borders.items.keys.sort
    b = @ole.borders.items[7]
    assert_equal 1, b.writes[:LineStyle]
    assert_equal(-4138, b.writes[:Weight])
    assert_equal WineOLE::MSOffice::Color['#999999'], b.writes[:Color]
  end

  def test_every_border_style
    {hairline: [1, 1], thin: [1, 2], medium: [1, -4138], thick: [1, 4],
     dash: [-4115, 2], dot: [-4118, 2]}.each do |sym, (line, weight)|
      setup
      apply(border: {edges: :bottom, style: sym})
      b = @ole.borders.items[9]
      assert_equal line, b.writes[:LineStyle], "style: #{sym} line"
      assert_equal weight, b.writes[:Weight], "style: #{sym} weight"
    end
  end

  def test_the_hash_form_defaults_to_thin
    apply(border: {edges: :bottom})
    assert_equal 2, @ole.borders.items[9].writes[:Weight]
  end

  # :none has no line to colour: the colour is still validated (a bad one
  # raises), but there is deliberately nothing left to write it to.
  def test_none_with_a_colour_validates_but_never_writes_the_colour
    apply(border: {edges: :all, style: :none, color: '#FF0000'})
    assert_equal(-4142, @ole.borders.writes[:LineStyle])
    refute @ole.borders.writes.key?(:Weight)
    refute @ole.borders.writes.key?(:Color),
      'style: :none has no line to colour, so Color must never be written'
  end

  # The error for a bad edge says "an array of those", and "those" includes
  # :all and :outline -- so both must actually work inside an array.
  def test_all_expands_inside_an_array
    apply(border: [:all])
    assert_equal 1, @ole.borders.writes[:LineStyle]
    assert_equal 0, @ole.borders.fetches, 'an array containing only :all still takes the bulk path'
  end

  def test_outline_expands_inside_an_array
    apply(border: [:outline])
    assert_equal [7, 8, 9, 10].sort, @ole.borders.items.keys.sort
  end

  # Outline plus a specific inside edge -- impossible before shorthands
  # expanded inside an array, and a reasonable thing to ask for.
  def test_outline_can_be_combined_with_another_edge_in_an_array
    apply(border: [:outline, :inside_h])
    assert_equal [7, 8, 9, 10, 12].sort, @ole.borders.items.keys.sort
  end

  # An empty edge list writes nothing, so it must not even fetch Borders.
  def test_an_empty_edge_list_fetches_nothing
    apply(border: [])
    assert_equal 0, @ole.borders_fetches
  end

  # [:top, :top] must fetch and write Item(8) once, not twice.
  def test_duplicate_edges_are_written_once
    apply(border: [:top, :top])
    assert_equal 1, @ole.borders.fetches, 'a duplicated edge must be deduplicated before writing'
  end

  def test_borders_is_fetched_once
    apply(border: :all)
    assert_equal 1, @ole.borders_fetches,
      'each `ole.Borders` is a round trip -- fetch it once and Item() off it'
  end

  def test_an_unknown_edge_is_refused
    err = assert_raises(ArgumentError) { apply(border: :diagonal) }
    assert_match(/border/, err.message)
    assert_match(/outline/, err.message)
  end

  def test_an_unknown_border_style_is_refused
    err = assert_raises(ArgumentError) { apply(border: {edges: :all, style: :squiggly}) }
    assert_match(/style/, err.message)
    assert_match(/thin/, err.message)
  end

  def test_a_raw_colour_integer_in_a_border_is_refused
    assert_raises(ArgumentError) { apply(border: {edges: :all, color: 255}) }
  end

  def test_an_unknown_key_inside_the_border_hash_is_refused
    err = assert_raises(ArgumentError) { apply(border: {edges: :all, widht: 2}) }
    assert_match(/widht/, err.message)
  end

  # nil means "no value for this" everywhere else in this module -- an
  # explicit style: nil inside the border hash must fall back to :thin
  # rather than override the default with a value that then fails to
  # validate.
  def test_nil_inside_the_border_hash_falls_back_to_the_default
    apply(border: {edges: :all, style: nil})
    assert_equal 2, @ole.borders.writes[:Weight], 'style: nil must fall back to :thin'
  end

  def test_border_nil_means_not_specified
    apply(border: nil, bold: true)
    assert_equal 0, @ole.borders_fetches, 'border: nil must be treated as "not specified"'
    assert_equal true, @ole.font.writes[:Bold]
  end

  # Same discipline as every other key: a bad value leaves the range alone.
  def test_a_bad_border_is_caught_before_anything_is_written
    assert_raises(ArgumentError) { apply(bold: true, border: :diagonal) }
    assert_equal 0, @ole.borders_fetches
    assert_equal 0, @ole.font_fetches,
      'nothing may be written until the whole option hash has been validated'
  end

  # --- reject_unknown_keys runs before the nil drop ----------------------

  # A misspelled key whose value happens to be nil used to be silently
  # absorbed, because `apply` dropped nils before checking key names. Key
  # names must be checked against the hash as given.
  def test_a_misspelled_key_with_a_nil_value_is_still_refused
    err = assert_raises(ArgumentError) { apply(blod: nil) }
    assert_match(/blod/, err.message)
  end

  def test_a_correct_key_with_a_nil_value_is_still_a_no_op
    apply(bold: nil)
    assert_empty @ole.font.writes
    assert_equal 0, @ole.font_fetches
  end
end
