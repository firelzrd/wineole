require_relative 'color'

module WineOLE
  module MSOffice
    # The one place that knows how a format key maps onto a COM property.
    #
    # Keyword arguments rather than a chain, because formatting needs three
    # states and a chain has two: an absent key means "leave it alone", and
    # `false` means "explicitly turn it off". Those are different operations
    # and both are ordinary. The implementation therefore asks `key?` and
    # never leans on `opts[:bold]` being nil.
    #
    # Every number here was measured against a live Excel 11 rather than
    # recalled; see the spec's table.
    module Format
      UNDERLINE = {none: -4142, single: 2, double: -4119}.freeze
      ALIGN = {general: 1, left: -4131, center: -4108, right: -4152, justify: -4130}.freeze
      VALIGN = {top: -4160, center: -4108, bottom: -4107}.freeze
      private_constant :UNDERLINE, :ALIGN, :VALIGN

      XL_NONE = -4142
      XL_AUTOMATIC = -4105
      XL_GENERAL_FORMAT_NAME = 26
      private_constant :XL_NONE, :XL_AUTOMATIC, :XL_GENERAL_FORMAT_NAME

      FONT_KEYS = %i[bold italic underline size color].freeze
      INTERIOR_KEYS = %i[background].freeze
      RANGE_KEYS = %i[align valign wrap number_format].freeze
      KEYS = (FONT_KEYS + INTERIOR_KEYS + RANGE_KEYS + %i[border]).freeze
      private_constant :FONT_KEYS, :INTERIOR_KEYS, :RANGE_KEYS, :KEYS

      # One knob for the caller, two properties for Excel. Excel keeps the
      # line pattern (LineStyle) and its thickness (Weight) apart, which
      # means `:thin` and `:dash` live in different properties even though a
      # caller thinks of both as "what the line looks like". Measured pairs:
      BORDER_STYLE = {
        none:     {line: -4142, weight: nil},
        hairline: {line: 1,     weight: 1},
        thin:     {line: 1,     weight: 2},
        medium:   {line: 1,     weight: -4138},
        thick:    {line: 1,     weight: 4},
        dash:     {line: -4115, weight: 2},
        dot:      {line: -4118, weight: 2},
      }.freeze
      private_constant :BORDER_STYLE

      EDGE = {left: 7, top: 8, bottom: 9, right: 10, inside_v: 11, inside_h: 12}.freeze
      OUTLINE_EDGES = %i[left top bottom right].freeze
      ALL_EDGES = %i[left top bottom right inside_v inside_h].freeze
      BORDER_HASH_KEYS = %i[edges style color].freeze
      private_constant :EDGE, :OUTLINE_EDGES, :ALL_EDGES, :BORDER_HASH_KEYS

      # Two passes on purpose. Everything is validated and converted into the
      # values COM wants *before* the first write, so a bad key or a bad
      # value leaves the range exactly as it was. Validating as it went would
      # mean `format(bold: true, align: :middle)` raises with the range
      # already bold -- and a caller who sees an exception reasonably reads
      # it as "nothing happened".
      #
      # `translate` may read from COM (`:general` asks the Application for
      # the local name of the General format). A read changes nothing, so
      # the invariant that matters -- no write before validation finishes --
      # still holds.
      def self.apply(ole, opts)
        reject_unknown_keys(opts)
        opts = opts.reject { |_k, v| v.nil? } # nil means "not specified"
        font, interior, range, border = translate(ole, opts)

        # Each `ole.Font` is its own round trip, so it is fetched once and
        # only when there is something to write to it.
        write_to(ole.Font, font) unless font.empty?
        write_to(ole.Interior, interior) unless interior.empty?
        write_to(ole, range)
        write_border(ole, border) unless border.nil?
        nil
      end

      def self.write_to(target, assignments)
        assignments.each { |setter, value| target.public_send(setter, value) }
      end
      private_class_method :write_to

      # Returns four values: three lists of [setter, value] -- for Font, for
      # Interior and for the Range itself -- plus a border plan (or nil) for
      # `write_border`, which does not fit the [setter, value] shape because
      # it may need both a bulk assignment and a per-edge loop. Raises
      # rather than returning anything partial.
      def self.translate(ole, opts)
        font = []
        interior = []
        range = []

        font << [:Bold=, boolean(opts[:bold], :bold)] if opts.key?(:bold)
        font << [:Italic=, boolean(opts[:italic], :italic)] if opts.key?(:italic)
        font << [:Underline=, underline(opts[:underline])] if opts.key?(:underline)
        font << [:Size=, size(opts[:size])] if opts.key?(:size)
        if opts.key?(:color)
          font << if opts[:color] == false
                    [:ColorIndex=, XL_AUTOMATIC]
                  else
                    [:Color=, colour(opts[:color], :color)]
                  end
        end

        if opts.key?(:background)
          interior << if opts[:background] == false
                        # Not `Color = white`: measured, a cleared cell and a
                        # white-painted cell both report Color 16777215, but
                        # the painted one keeps ColorIndex 2 and Pattern 1 --
                        # it is still filled, prints as a fill, and hides
                        # gridlines.
                        [:ColorIndex=, XL_NONE]
                      else
                        [:Color=, colour(opts[:background], :background)]
                      end
        end

        range << [:HorizontalAlignment=, fetch(ALIGN, opts[:align], :align)] if opts.key?(:align)
        range << [:VerticalAlignment=, fetch(VALIGN, opts[:valign], :valign)] if opts.key?(:valign)
        range << [:WrapText=, boolean(opts[:wrap], :wrap)] if opts.key?(:wrap)
        if opts.key?(:number_format)
          range << [:NumberFormat=, number_format(ole, opts[:number_format])]
        end

        border = opts.key?(:border) ? translate_border(opts[:border]) : nil

        [font, interior, range, border]
      end
      private_class_method :translate

      # Validates and resolves everything; touches no COM.
      def self.translate_border(spec)
        spec = normalize_border(spec)
        style = fetch(BORDER_STYLE, spec[:style], 'border style')
        # Named line_colour, not colour: a local called `colour` would shadow
        # the method of that name for the rest of this body.
        line_colour = spec.key?(:color) ? colour(spec[:color], 'border color') : nil

        {indexes: expand_edges(spec[:edges]), line: style[:line],
         weight: style[:weight], colour: line_colour}
      end
      private_class_method :translate_border

      # `ole.Borders` is a round trip like Font and Interior: fetched once,
      # with Item() called off it -- except when every edge is being set, in
      # which case an assignment straight to the Borders collection replaces
      # the whole per-edge loop.
      #
      # Measured against a live Excel on a multi-cell range, an assignment on
      # Borders itself reaches all six edges -- including inside_v and
      # inside_h -- in one COM call each: LineStyle, Weight and Color each
      # touch all six for the price of one round trip. 2.3 ms against 10.0 ms
      # for the per-edge Item() loop below. That is exactly why it must never
      # be used for :outline: it would silently draw the inside edges too.
      # Keyed off the resolved index set rather than the :all symbol, so an
      # explicit list of all six edges gets the fast path as well.
      ALL_EDGE_INDEXES = ALL_EDGES.map { |name| EDGE.fetch(name) }.sort.freeze
      private_constant :ALL_EDGE_INDEXES

      def self.write_border(ole, plan)
        indexes = plan[:indexes].uniq
        return if indexes.empty?

        borders = ole.Borders

        if indexes.sort == ALL_EDGE_INDEXES
          borders.LineStyle = plan[:line]
          return if plan[:weight].nil?

          borders.Weight = plan[:weight]
          borders.Color = plan[:colour] unless plan[:colour].nil?
          return
        end

        indexes.each do |index|
          edge = borders.Item(index)
          edge.LineStyle = plan[:line]
          # Nothing to weigh or colour when the line is being removed.
          next if plan[:weight].nil?

          edge.Weight = plan[:weight]
          edge.Color = plan[:colour] unless plan[:colour].nil?
        end
      end
      private_class_method :write_border

      def self.normalize_border(spec)
        hash = case spec
               when false     then {edges: :all, style: :none}
               when ::Symbol  then {edges: spec, style: :thin}
               when ::Array   then {edges: spec, style: :thin}
               when ::Hash    then spec
               else
                 raise ArgumentError,
                   'border: expected :all, :outline, an edge name, an array of edge ' \
                   "names, false, or a hash, got #{spec.inspect}"
               end
        # Checked against the hash as given, before nils are dropped -- same
        # order as `apply`'s top-level keys, so a misspelled key with a nil
        # value is still caught rather than silently absorbed.
        unknown = hash.keys - BORDER_HASH_KEYS
        unless unknown.empty?
          raise ArgumentError,
            "border: unknown key#{'s' if unknown.length > 1} " \
            "#{unknown.map(&:inspect).join(', ')} -- known keys are #{BORDER_HASH_KEYS.join(', ')}"
        end

        # nil means "not specified" everywhere else in this module (`apply`
        # drops nil top-level values before anything is validated); a caller
        # writing `style: nil` explicitly is asking for the same thing as
        # omitting `style`, not for an override that beats the default.
        {style: :thin}.merge(hash.compact)
      end
      private_class_method :normalize_border

      def self.expand_edges(edges)
        names = case edges
                # Only a hash form reaches here without edges: having been
                # filled in -- the shorthands (:all, an edge name, an array,
                # false) all set it themselves in normalize_border. Naming
                # the missing key beats bad_edges_message's "got nil", which
                # reads as a value the caller never wrote.
                when nil
                  raise ArgumentError,
                    'border: a hash needs an edges: key, e.g. ' \
                    'border: {edges: :all, style: :thick}'
                when :all      then ALL_EDGES
                when :outline  then OUTLINE_EDGES
                when ::Symbol  then [edges]
                # :all and :outline expand here too, not just on their own --
                # the error message below promises "an array of those", and
                # "those" includes the shorthands. This also makes something
                # like [:outline, :inside_h] expressible.
                when ::Array
                  edges.flat_map do |edge|
                    case edge
                    when :all     then ALL_EDGES
                    when :outline then OUTLINE_EDGES
                    else [edge]
                    end
                  end
                else raise ArgumentError, bad_edges_message(edges)
                end
        names.map do |name|
          EDGE.fetch(name) { raise ArgumentError, bad_edges_message(name) }
        end
      end
      private_class_method :expand_edges

      # One message for both failures, and it names the shorthands as well as
      # the edges -- someone who typed :diagonal needs to learn that :all and
      # :outline exist, which a bare list of EDGE's keys would not tell them.
      def self.bad_edges_message(value)
        "border: expected :all, :outline, one of " \
        "#{EDGE.keys.map(&:inspect).join(', ')}, or an array of those, " \
        "got #{value.inspect}"
      end
      private_class_method :bad_edges_message

      # Before any COM call, so a typo leaves the sheet exactly as it was
      # rather than half-formatted.
      def self.reject_unknown_keys(opts)
        unknown = opts.keys - KEYS
        return if unknown.empty?

        raise ArgumentError,
          "unknown format key#{'s' if unknown.length > 1} " \
          "#{unknown.map(&:inspect).join(', ')} -- known keys are #{KEYS.join(', ')}"
      end
      private_class_method :reject_unknown_keys

      # A nil never reaches here -- `apply` drops nil values first, because
      # nil means "I have no value for this", which is the same thing as not
      # passing the key. That matters: measured, assigning nil to a COM
      # boolean property sets it to *false*, so a nil reaching COM would
      # silently un-bold a range whose caller simply did not know.
      def self.boolean(value, key)
        return value if value == true || value == false

        raise ArgumentError,
          "#{key}: expected true or false, got #{value.inspect}. " \
          'Omit the key entirely to leave this attribute alone'
      end
      private_class_method :boolean

      # Excel's own font size range. Outside it, COM fails with a message
      # about the Font class rather than about the number.
      SIZE_RANGE = (1..409).freeze
      private_constant :SIZE_RANGE

      def self.size(value)
        unless value.is_a?(::Numeric) && SIZE_RANGE.cover?(value)
          raise ArgumentError,
            "size: expected a number in 1..409 (Excel's own range), got #{value.inspect}"
        end

        value
      end
      private_class_method :size

      def self.underline(value)
        case value
        when true  then UNDERLINE.fetch(:single)
        when false then UNDERLINE.fetch(:none)
        else fetch(UNDERLINE, value, :underline)
        end
      end
      private_class_method :underline

      # A raw COM colour integer is ambiguous here: 255 could mean the
      # caller's #0000FF or Excel's own value for red. Refuse rather than
      # guess -- the same stance `write` takes on a wrong-shaped array.
      def self.colour(value, key)
        if value.is_a?(::Numeric)
          raise ArgumentError,
            "#{key}: expected '#RRGGBB', got the number #{value.inspect}. " \
            'A raw COM colour is ambiguous here -- pass a hex string, or use ' \
            'WineOLE::MSOffice::Color[...] with .ole to reach COM directly'
        end

        begin
          Color[value]
        rescue ArgumentError => e
          raise ArgumentError, "#{key}: #{e.message}"
        end
      end
      private_class_method :colour

      # 'General' is the one format code that cannot be written: measured, it
      # fails outright on a localized Excel, where the format has a
      # translated name instead -- and that translated spelling is not
      # portable either. Application.International(26) returns whichever
      # one this Excel wants.
      def self.number_format(ole, value)
        case value
        when :general then ole.Application.International(XL_GENERAL_FORMAT_NAME)
        when :text then '@'
        when ::String
          if value.casecmp?('General')
            ole.Application.International(XL_GENERAL_FORMAT_NAME)
          else
            value
          end
        else
          raise ArgumentError,
            "number_format: expected a format code string, :general or :text, " \
            "got #{value.inspect}"
        end
      end
      private_class_method :number_format

      def self.fetch(table, value, key)
        table.fetch(value) do
          raise ArgumentError,
            "#{key}: expected one of #{table.keys.map(&:inspect).join(', ')}, " \
            "got #{value.inspect}"
        end
      end
      private_class_method :fetch
    end
  end
end
