require 'English'
require_relative '../errors'

module WineOLE
  module MSOffice
    # Programmatic access to a workbook's VBA project is off by default, and
    # turning it on means writing a macro security setting in the registry.
    # This module can read and write it -- but nothing in the wrapper calls
    # the writing half. That is for `wineole-vba`, which a human runs, and
    # for tests, which need to exercise both sides of the switch.
    module VBA
      class Error < WineOLE::Error; end
      class AccessDenied < Error; end

      ACCESS_KEY = 'HKCU\Software\Microsoft\Office\11.0\Excel\Security'.freeze
      ACCESS_VALUE = 'AccessVBOM'.freeze
      CODEPAGE_KEY = 'HKLM\System\CurrentControlSet\Control\Nls\CodePage'.freeze

      # :enabled, :disabled, or :unset when the value is not there at all.
      def self.state
        raw = read(ACCESS_KEY, ACCESS_VALUE)
        return :unset if raw.nil?

        raw.to_i(16).zero? ? :disabled : :enabled
      end

      def self.enabled?
        state == :enabled
      end

      # Touching VBProject when access is refused gives 0x800A03EC and a
      # localized message -- the same HRESULT a rejected NumberFormat
      # gives, and neither identifies the condition. The registry is what
      # turns the refusal into advice.
      def self.denied!
        message =
          case state
          when :enabled
            'access to the VBA project was refused even though the registry ' \
            'has it enabled -- Excel reads that setting when it starts, so ' \
            'restart Excel if it was switched on while this instance was running'
          else
            'access to the VBA project is disabled. Run `wineole-vba enable`, ' \
            'then restart Excel -- it reads the setting at startup'
          end
        raise AccessDenied, message
      end

      def self.enable!
        write(ACCESS_KEY, ACCESS_VALUE, '1')
      end

      def self.disable!
        write(ACCESS_KEY, ACCESS_VALUE, '0')
      end

      # The Windows ANSI codepage of this prefix, as a name Ruby knows.
      # Never hardcoded: VBA source files are written and read in whatever
      # this is, and it is not CP932 everywhere.
      #
      # Memoized because it costs a `wine reg` subprocess -- measured at
      # 328 ms on this host -- and a machine's ANSI codepage does not change
      # while a process runs. Without this an import pays it twice and a
      # non-ASCII refusal pays it three times.
      def self.codepage
        @codepage ||= read_codepage
      end

      # For tests, which swap the codepage to exercise both sides of it.
      def self.forget_codepage
        @codepage = nil
      end

      def self.read_codepage
        raw = read(CODEPAGE_KEY, 'ACP')
        raise Error, "could not read the ANSI codepage (ACP) from #{CODEPAGE_KEY}" if raw.nil?

        name = "CP#{raw}"
        begin
          ::Encoding.find(name)
        rescue ArgumentError
          raise Error, "the registry reports ANSI codepage #{raw.inspect}, which Ruby does not know"
        end
        name
      end

      # One explanation, used by both paths that hand text to Excel. They
      # are bound by the same codepage and fail the same way, so the caller
      # should not be able to tell from the message which one they hit --
      # only which character stopped it.
      #
      # `where` says what the text was, because the way out differs: code
      # given as a string can be rewritten with ChrW(), a file has to be
      # edited.
      def self.unrepresentable!(char, where)
        raise ArgumentError,
          "#{where} contains #{char.inspect}, which the system codepage (#{codepage}) " \
          "cannot represent. Excel stores a module's text in that codepage, so the " \
          'character would be silently replaced rather than stored. Rewrite it with ' \
          'Chr()/ChrW() escapes, which are built at run time and are not bound by the ' \
          'codepage the source text is'
      end

      # BOMs, longest first -- FF FE 00 00 is UTF-32LE and also starts with
      # the UTF-16LE BOM, so a shorter match must never be tried first.
      BOMS = [
        ["\x00\x00\xFE\xFF".b, 'UTF-32BE'],
        ["\xFF\xFE\x00\x00".b, 'UTF-32LE'],
        ["\xEF\xBB\xBF".b,      'UTF-8'],
        ["\xFE\xFF".b,           'UTF-16BE'],
        ["\xFF\xFE".b,           'UTF-16LE']
      ].freeze

      # What encoding a VBA source file should be read as. Three rules, and
      # every one of them decides on evidence rather than on a guess -- an
      # encoding a heuristic merely finds likely is exactly the "succeeded,
      # returned a value, the value is wrong" failure this wrapper exists to
      # remove.
      #
      #   1. A BOM is conclusive. Follow it.
      #   2. Bytes that are not valid UTF-8 PROVE the file is not UTF-8, so
      #      read it as the ANSI codepage -- which is what Excel's own
      #      Export writes, and what every .bas from a Windows toolchain is.
      #   3. Otherwise UTF-8.
      #
      # Rule 2 is the one that earns its place, and the direction matters:
      # measured on this host, a CP932 file read as UTF-8 is invalid 95.07%
      # of the time at ONE non-ASCII character and 99.99% by five, so real
      # codepage files land here almost without exception. The reverse does
      # not hold -- UTF-8 bytes read as CP932 come out VALID from two
      # characters on, silently wrong. That asymmetry is why UTF-8 is the
      # fallback in rule 3 and the codepage is never the default: guessing
      # UTF-8 and being wrong is loud, guessing the codepage and being wrong
      # is silent.
      #
      # What is left is a file in the codepage whose bytes happen to be
      # valid UTF-8 -- undecidable, by construction, for anyone. Pass
      # `encoding:` to skip all of this when you already know.
      def self.detect_encoding(path)
        head = ::File.binread(path, 4).to_s
        BOMS.each { |bytes, name| return name if head.start_with?(bytes) }

        ::File.binread(path).force_encoding('UTF-8').valid_encoding? ? 'UTF-8' : codepage
      end

      def self.read(key, value)
        out, ok = run_reg(['query', key, '/v', value])
        return nil unless ok

        # "    NAME    TYPE    VALUE", indented, and wine leaves a CR on the
        # end. Stripping each line before splitting is what makes both go
        # away: without it the leading spaces produce an empty first field
        # and the value comes back as "REG_SZ    932\r\n".
        #
        # The name is matched as a whole field, not as a substring of the
        # line -- otherwise a value name that happens to appear inside
        # another value's data would match, and it would happen silently.
        matches = out.lines.filter_map do |l|
          parts = l.strip.split(/\s+/, 3)
          parts if parts.length == 3 && parts[0] == value
        end

        if matches.length > 1
          raise Error,
                "found #{matches.length} lines naming #{value} in `wine reg query #{key}` " \
                "output, expected exactly one: #{out.inspect}"
        end

        if matches.empty?
          raise Error,
                "the command succeeded but no line naming #{value} could be parsed from " \
                "`wine reg query #{key}` output: #{out.inspect}"
        end

        matches.first[2]
      end
      private_class_method :read

      def self.write(key, value, data)
        _out, ok = run_reg(['add', key, '/v', value, '/t', 'REG_DWORD', '/d', data, '/f'])
        ok
      end
      private_class_method :write

      # The only place that shells out, and deliberately not private: it is
      # the seam the tests replace so that no test touches a real registry.
      #
      # Exit status is the only trustworthy signal. `reg` writes both its
      # success message and its not-found message to stdout, in the system
      # language, and wine writes unrelated `fixme:` lines to stderr.
      def self.run_reg(args)
        out = IO.popen(['wine', 'reg', *args], err: File::NULL, &:read)
        [out.to_s, $CHILD_STATUS&.success? || false]
      rescue SystemCallError, IOError
        ['', false]
      end
    end
  end
end
