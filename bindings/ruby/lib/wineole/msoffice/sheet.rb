require_relative '../proxy'
require_relative 'passthrough'
require_relative 'address'
require_relative 'range'
require_relative 'vba_api'
require_relative 'controls'

module WineOLE
  module MSOffice
    # Wraps a COM Worksheet. `[]`/`[]=` address it through the same `Address`
    # parser the whole wrapper uses, so `sheet['A1:B2']` and
    # `sheet[row, col]` both hand back a `Range`; everything else falls
    # through to COM.
    class Sheet
      include Passthrough

      def initialize(proxy, version:)
        @ole = proxy
        @version = version
      end

      def [](*keys)
        range_for(keys)
      end

      # Delegates to Range#write, never #fill: write raises on a value that
      # does not fit the range rather than replicating or padding it, which
      # is the behaviour an assignment through []= should have.
      def []=(*args)
        value = args.pop
        range_for(args).write(value)
      end

      # This worksheet's VBA surface. Blocks only -- a worksheet's module
      # cannot be created or deleted, so SheetVBA has no component methods
      # rather than having ones that always fail.
      def vba
        @vba ||= SheetVBA.new(@ole)
      end

      # The Forms-toolbar controls on this sheet: `form_controls.add(:button,
      # name: 'Go', at: 'B2')`. See FormControls for what they can and
      # cannot do; `activex` is the family whose events reach Ruby.
      #
      # Name note (M0, measured against Excel 11): `form_controls` and
      # `activex` are both free on Worksheet.
      def form_controls
        @form_controls ||= FormControls.new(self)
      end

      # The ActiveX controls on this sheet, MSForms or any registered
      # ProgID: `activex.add(:command_button, name: 'Go', at: 'B2')`.
      def activex
        @activex ||= ActiveXControls.new(self)
      end

      private

      def range_for(keys)
        return Range.new(@ole.Cells(*keys)) if cell_reference?(keys)

        unless keys.length == 1 && keys.first.is_a?(::String)
          raise ArgumentError, "unsupported sheet index #{keys.inspect}"
        end

        str = keys.first
        addr = Address.parse(str, @version)

        # An address that parses but stops at a sheet or a book (or does
        # not parse at all) is a lookup, not something with cells to read
        # or write -- see the spec on why `xl["Sheet1!"] = 0` filling a
        # whole sheet is a hazard rather than a convenience (Spec §4.5).
        if addr.nil? || !addr.range?
          raise ArgumentError, "#{str.inspect} has no range"
        end

        # This object is one sheet. An address that names a different
        # workbook or worksheet would silently reach past it -- writing to
        # the sheet the caller *named* instead of the one they *have* is
        # exactly the class of silent wrong-target write this wrapper
        # exists to prevent.
        if named?(addr.workbook) || named?(addr.worksheet)
          raise ArgumentError,
            "#{str.inspect} names another workbook or worksheet; a Sheet " \
            'addresses its own cells only -- reach another sheet through ' \
            'the Excel object that owns it'
        end

        Range.new(@ole.Range(addr.range))
      end

      def cell_reference?(keys)
        keys.length == 2 && keys.all? { |k| k.is_a?(::Integer) }
      end

      def named?(part)
        !part.nil? && !part.empty?
      end
    end
  end
end
