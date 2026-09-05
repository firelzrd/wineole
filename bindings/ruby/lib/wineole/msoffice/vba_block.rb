require_relative 'vba'

module WineOLE
  module MSOffice
    # A named span of code inside a VBA module, delimited by comment markers.
    #
    # The wrapper owns the span, never the module. A module may be one the
    # caller wrote in, or one that cannot be deleted at all (ThisWorkbook, a
    # worksheet), so "replace the module" is not available and would be too
    # blunt even where it is.
    module VBABlock
      NAME = /\A[A-Za-z0-9_-]+\z/.freeze

      # Matches any wineole marker line, open or close, for any name -- not
      # just the one being written. Used to refuse a payload that would be
      # indistinguishable from the wrapper's own delimiters.
      MARKER_LINE = /\A'<\/?wineole:[A-Za-z0-9_-]+>\z/.freeze

      def self.open_marker(name)  = "'<wineole:#{name}>"
      def self.close_marker(name) = "'</wineole:#{name}>"

      # Replaces the block of this name if it is there, adds it if not.
      def self.write(code_module, name, code)
        check_name(name)
        check_representable(code)
        check_payload(code)
        remove(code_module, name)
        code_module.AddFromString(
          "#{open_marker(name)}\n#{code.chomp}\n#{close_marker(name)}\n"
        )
        nil
      end

      # False when there was nothing of this name to remove. Otherwise the
      # module's remaining lines, as an Array -- so a caller that needs to
      # know whether the module is now blank does not have to fetch the
      # whole body a second time to ask.
      #
      # Block names match case-insensitively, as VBA identifiers do: two
      # blocks named `main` and `Main` would hold procedures that collide,
      # and VBA answers "Ambiguous name detected" for the whole module from
      # then on. So `Main` replaces (and removes) a block written as `main`.
      def self.remove(code_module, name)
        check_name(name)
        lines = body(code_module)
        return false if lines.empty?

        first = lines.index { |l| l.strip.casecmp?(open_marker(name)) }
        last = lines.index { |l| l.strip.casecmp?(close_marker(name)) }

        if first.nil?
          if last
            raise ArgumentError,
              "the #{name.inspect} block in this module has a closing marker " \
              'with no matching opening one -- the module is already ' \
              'corrupted; refusing to guess what to remove'
          end
          return false
        end

        if last.nil? || last < first
          raise ArgumentError,
            "the #{name.inspect} block in this module has no closing marker -- " \
            'refusing to guess where it ends'
        end

        code_module.DeleteLines(first + 1, last - first + 1)
        lines[0...first] + lines[(last + 1)..]
      end

      # Nothing but whitespace. Not CountOfLines == 0: a module emptied of
      # its blocks still reports the newlines that held them.
      def self.blank?(code_module)
        blank_lines?(body(code_module))
      end

      # The same emptiness rule #blank? uses, applied to lines the caller
      # already has (typically the Array #remove just handed back) instead
      # of fetching the body again.
      def self.blank_lines?(lines)
        lines.all? { |l| l.strip.empty? }
      end

      # One round trip, whatever the module's length. Excel reports 0 lines
      # for a module never written to, and Lines(1, 0) is not a legal call.
      def self.body(code_module)
        count = code_module.CountOfLines.to_i
        return [] if count.zero?

        code_module.Lines(1, count).split(/\r?\n/)
      end
      private_class_method :body

      # A module's text is held in the system ANSI codepage, not Unicode --
      # measured, not assumed: on a CP932 host `café` comes back `cafe`,
      # `✓` comes back `?`, and simplified Chinese comes back part `?`.
      # Japanese survives only because CP932 can represent it, which is why
      # an earlier measurement using Japanese alone concluded, wrongly, that
      # this path carried Unicode.
      #
      # So this path is bound by exactly the same codepage as import_vba,
      # and gets the same rule: refuse rather than substitute. Silently
      # dropping an accent is the failure this whole phase exists to remove.
      #
      # ASCII skips the check, which is almost every call -- resolving the
      # codepage costs a `wine reg` invocation.
      def self.check_representable(code)
        return if code.ascii_only?

        code.encode(VBA.codepage)
      rescue ::Encoding::UndefinedConversionError => e
        VBA.unrepresentable!(e.error_char, 'this code')
      end
      private_class_method :check_representable

      def self.check_name(name)
        return if name.is_a?(::String) && name.match?(NAME)

        raise ArgumentError,
          "a block name must match #{NAME.inspect} -- it goes inside a VBA " \
          "comment marker, so it cannot contain spaces, '>' or newlines. " \
          "Got #{name.inspect}"
      end
      private_class_method :check_name

      # A code body containing a line that is itself a wineole marker would
      # be indistinguishable from a real one once written: #remove would
      # find the caller's accidental marker instead of its own, delete up to
      # the wrong place, and leave the rest as permanent garbage. The
      # wrapper controls what it writes, so refusing the payload up front is
      # what keeps that state from ever existing.
      def self.check_payload(code)
        offending = code.split(/\r?\n/).find { |line| line.strip.match?(MARKER_LINE) }
        return if offending.nil?

        raise ArgumentError,
          "the code being written contains a line that is itself a wineole " \
          "marker (#{offending.strip.inspect}) -- this cannot be told apart " \
          'from the wrapper\'s own markers, so it is refused'
      end
      private_class_method :check_payload
    end
  end
end
