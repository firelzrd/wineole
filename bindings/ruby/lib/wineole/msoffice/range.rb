require_relative '../proxy'
require_relative 'passthrough'
require_relative 'format'

module WineOLE
  module MSOffice
    # Wraps a COM Range. Adds exactly four methods plus `ole`; everything
    # else falls through to COM.
    #
    # The restraint is deliberate. COM's Range already exposes Rows, Columns,
    # Cells, Item, Areas, Find, Sort, Merge, Table and many more, and COM
    # resolves names case-insensitively -- so every lowercase single word this
    # class defines covers a COM call that already worked. `to_a`, `write`,
    # `fill`, `format` and `ole` were each checked against a live Range and
    # found absent.
    class Range
      include Passthrough

      def initialize(proxy)
        @ole = proxy
      end

      # Always a two-dimensional Array, whatever the range's size.
      #
      # Excel's own `Value` returns a bare scalar for a one-cell range --
      # consistently, whether it was addressed as "A1" or "A1:A1" -- so
      # generic code that does not know the size has to branch. This is that
      # branch, written once.
      #
      # Returns a plain Array on purpose: Ruby's own vocabulary then covers
      # the rest (`rows.transpose` for columns, `rows.flatten` for values),
      # so this class does not have to grow `each_row`, `each_column` or
      # `flatten` and risk shadowing more COM members.
      def to_a
        v = @ole.Value
        v.is_a?(::Array) ? v : [[v]]
      end

      # Write, refusing anything that does not fit.
      #
      # Excel's own assignment corrupts silently in three ways (all measured):
      # a flat array written to a column replicates its first element down
      # every cell, too few values leave #N/A behind, and too many are
      # truncated. None of those raise, and none are visible without looking
      # at the sheet.
      def write(value)
        @ole.Value = shaped(value)
      end

      # Write, adapting the value to the range: replicate along a dimension
      # the argument does not have, truncate or pad along one it does.
      #
      # Total -- every input has a defined result, no exceptions -- but note
      # that it reproduces Excel's own column trap by construction: a flat
      # array is a row, so filling an Nx1 column with [1,2,3] puts 1 in every
      # cell. That is why `write` and not `fill` is what `sheet[addr] = x`
      # uses; reach for this one deliberately.
      def fill(value)
        nrows = row_count
        ncols = column_count
        @ole.Value =
          case value
          when ::Array
            if value.first.is_a?(::Array)
              (0...nrows).map { |r| row = value[r] || []; (0...ncols).map { |c| row[c] } }
            else
              one = (0...ncols).map { |c| value[c] }
              ::Array.new(nrows) { one.dup }
            end
          else
            ::Array.new(nrows) { ::Array.new(ncols) { value } }
          end
      end

      # Apply formatting. Keys are documented on WineOLE::MSOffice::Format.
      #
      # An absent key leaves that attribute alone; `false` turns it off.
      # That third state is why this takes keyword arguments rather than
      # being a chain of verbs -- and it keeps this class's additions to one
      # name, which matters because COM resolves names case-insensitively
      # and every lowercase word here covers a COM member of the same
      # spelling. `format` was measured free on a live Range.
      #
      # Note for anyone editing this class: defining `format` shadows
      # Kernel#format inside these instance methods, so string formatting
      # here must be written as `::Kernel.format(...)`.
      def format(**opts)
        Format.apply(@ole, opts)
        self
      end

      private

      def row_count
        @ole.Rows.Count
      end

      def column_count
        @ole.Columns.Count
      end

      def shaped(value)
        return value unless value.is_a?(::Array)

        nrows = row_count
        ncols = column_count

        if value.first.is_a?(::Array)
          unless value.all? { |r| r.is_a?(::Array) }
            raise ArgumentError,
              "range is #{nrows}x#{ncols}, but the value mixes rows and scalars"
          end
          widths = value.map(&:length).uniq
          if widths.length > 1
            raise ArgumentError,
              "range is #{nrows}x#{ncols}, but the value has ragged rows " \
              "(#{widths.join(', ')} elements)"
          end
          if value.any? { |r| r.any? { |c| c.is_a?(::Array) } }
            raise ArgumentError,
              "range is #{nrows}x#{ncols}, but the value nests more than two deep"
          end
          unless value.length == nrows && widths.first == ncols
            raise ArgumentError,
              "range is #{nrows}x#{ncols}, but the value is " \
              "#{value.length}x#{widths.first}"
          end
          value
        else
          if value.any? { |v| v.is_a?(::Array) } # rubocop:disable Style/IfInsideElse
            raise ArgumentError,
              "range is #{nrows}x#{ncols}, but the value mixes scalars and rows"
          end
          if nrows > 1 && ncols > 1
            raise ArgumentError,
              "range is #{nrows}x#{ncols}; a flat array only fits a single " \
              'row or column -- pass rows, or use fill'
          end
          expected = nrows > 1 ? nrows : ncols
          unless value.length == expected
            raise ArgumentError,
              "range is #{nrows}x#{ncols}, but the value has #{value.length} elements"
          end
          nrows > 1 ? value.map { |v| [v] } : [value]
        end
      end
    end
  end
end
