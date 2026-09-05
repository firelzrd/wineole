module WineOLE
  module MSOffice
    # '#RRGGBB' <-> Excel's colour integer.
    #
    # Excel stores a colour as BGR, not RGB: measured by making Excel name
    # the colour itself, `Interior.Color = 255` reports ColorIndex 3 (red)
    # and `= 0xFF0000` reports ColorIndex 5 (blue). So the obvious
    # `'#FF0000'.delete('#').to_i(16)` produces blue, silently.
    #
    # Public, and returning a plain Integer, on purpose. The wrapper's own
    # keys are not the only place a colour is written -- the passthrough
    # reaches Interior.Color, Font.Color, Borders.Color, Tab.Color and
    # anything else COM has, and that surface is unbounded by design. One
    # function that returns an Integer covers all of it without the protocol
    # or Proxy#encode knowing anything about colours.
    module Color
      HEX = /\A#?(?:\h{3}|\h{6})\z/.freeze
      private_constant :HEX

      # '#RRGGBB' | '#RGB' | [r, g, b] -> Excel's integer.
      def self.[](value)
        r, g, b = rgb(value)
        r | (g << 8) | (b << 16)
      end

      # Excel's integer (or the Float it actually hands back) -> '#RRGGBB'.
      def self.to_hex(value)
        unless value.is_a?(::Numeric)
          raise ArgumentError, "expected a number from COM, got #{value.inspect}"
        end

        # Checked before converting: Float::INFINITY#to_i raises
        # FloatDomainError and Complex#to_i raises RangeError, neither of
        # which is this module's ArgumentError. Range#cover? (as
        # Format.size uses via SIZE_RANGE) rather than #between?, because
        # Complex does not include Comparable -- #between? raises
        # NoMethodError on it outright, where #cover? just answers false.
        unless (0..0xFFFFFF).cover?(value)
          raise ArgumentError,
            "expected a colour in 0..0xFFFFFF, got #{value.inspect}. " \
            "Excel's ColorIndex is a different property from Color -- its " \
            'values (-4105 automatic, -4142 none) are not colours and cannot ' \
            'be converted here'
        end
        n = value.to_i
        ::Kernel.format('#%02X%02X%02X', n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF)
      end

      def self.rgb(value)
        case value
        when ::String then from_hex(value)
        when ::Array  then from_array(value)
        else
          raise ArgumentError,
            "expected '#RRGGBB', '#RGB' or [r, g, b], got #{value.inspect}"
        end
      end
      private_class_method :rgb

      def self.from_hex(value)
        # Checked before parsing: String#to_i(16) reads garbage as 0 rather
        # than complaining, so '#GGGGGG' would silently become black.
        unless value.match?(HEX)
          raise ArgumentError,
            "expected '#RRGGBB' or '#RGB', got #{value.inspect}"
        end

        s = value.delete_prefix('#')
        s = s.chars.map { |c| c * 2 }.join if s.length == 3
        [s[0, 2], s[2, 2], s[4, 2]].map { |h| h.to_i(16) }
      end
      private_class_method :from_hex

      def self.from_array(value)
        ok = value.length == 3 &&
             value.all? { |c| c.is_a?(::Integer) && (0..255).cover?(c) }
        unless ok
          raise ArgumentError,
            "expected [r, g, b] with three integers in 0..255, got #{value.inspect}"
        end

        value
      end
      private_class_method :from_array
    end
  end
end
