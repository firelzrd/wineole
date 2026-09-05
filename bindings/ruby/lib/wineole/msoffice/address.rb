module WineOLE
  module MSOffice
    # Parses the addressing DSL this wrapper inherits from msoffice.rb:
    #
    #   "[Book1]Sheet1!A1:B2"   "[:new]"   ":new!"   ":first!A1"   "A1:B2"
    #
    # Touches no COM: it is a pure string parser, so it can be exercised
    # without Excel running.
    #
    # The patterns come from msoffice.rb, which encodes Excel's actual grid
    # limits (IV/65536 for Excel 11, XFD/1048576 for 12+). Both column
    # patterns were checked against every column up to their limit and are
    # exact -- and they are the harder half.
    #
    # The two row patterns needed a fix, carried here. Every nested digit
    # range started at 1 where it should start at 0: the leading digit's
    # "no leading zero" rule was carried down to positions the digits above
    # had already constrained. That dropped row 65530 outright, and 11,111
    # rows of the 1048576 grid. Both are exhaustively verified now, and
    # test_every_row_in_the_grid_parses is what keeps them that way.
    class Address
      PTN_A_IV = /[A-H]?[A-Z]|I[A-V]/i
      PTN_ABS_A_IV = /\$?#{PTN_A_IV}/
      PTN_1_65536 = /[1-9]\d{,3}|[1-5]\d{4}|6(?:[0-4]\d{3}|5(?:[0-4]\d{2}|5(?:[0-2]\d|3[0-6])))/
      PTN_ABS_1_65536 = /\$?#{PTN_1_65536}/
      PTN_ABS_A1_IV65536 = /#{PTN_ABS_A_IV}#{PTN_ABS_1_65536}/
      PTN_RANGE_LOCAL_XL11 = /(?<range>(?:#{PTN_ABS_A1_IV65536}:)?#{PTN_ABS_A1_IV65536}|#{PTN_ABS_A_IV}:#{PTN_ABS_A_IV}|#{PTN_ABS_1_65536}:#{PTN_ABS_1_65536})/
      PTN_A_XFD = /[A-W]?[A-Z]{1,2}|X(?:[A-E][A-Z]|F[A-D])/i
      PTN_ABS_A_XFD = /\$?#{PTN_A_XFD}/
      PTN_1_1048576 = /[1-9]\d{,5}|10(?:[0-3]\d{4}|4(?:[0-7]\d{3}|8(?:[0-4]\d{2}|5(?:[0-6]\d|7[0-6]))))/
      PTN_ABS_1_1048576 = /\$?#{PTN_1_1048576}/
      PTN_ABS_A1_XFD1048576 = /#{PTN_ABS_A_XFD}#{PTN_ABS_1_1048576}/
      PTN_RANGE_LOCAL_XL12 = /(?<range>(?:#{PTN_ABS_A1_XFD1048576}:)?#{PTN_ABS_A1_XFD1048576}|#{PTN_ABS_A_XFD}:#{PTN_ABS_A_XFD}|#{PTN_ABS_1_1048576}:#{PTN_ABS_1_1048576})/
      PTN_WORKBOOK = /\[(?<workbook>(?i::new)|[^\[\]]*)\]/
      PTN_WORKBOOK_WORKSHEET = /(?<worksheet_quote>'?)#{PTN_WORKBOOK}?(?<worksheet>(?i::(?:new|first|last))|(?:[^\[\]\\\:\*\'][^\[\]\\\:\*]*)?)\k<worksheet_quote>!/
      PTN_XL11 = /\A(?:#{PTN_WORKBOOK}|#{PTN_WORKBOOK_WORKSHEET}?#{PTN_RANGE_LOCAL_XL11}?)\z/
      PTN_XL12 = /\A(?:#{PTN_WORKBOOK}|#{PTN_WORKBOOK_WORKSHEET}?#{PTN_RANGE_LOCAL_XL12}?)\z/

      attr_reader :workbook, :worksheet, :range

      # Returns nil when the string is not an address at all, so a caller can
      # fall back to treating it as a raw sheet name -- which is what
      # msoffice.rb did.
      def self.parse(str, excel_version)
        pattern = excel_version.to_f >= 12 ? PTN_XL12 : PTN_XL11
        m = pattern.match(str)
        return nil if m.nil? || m.to_s.empty?

        new(workbook: m[:workbook], worksheet: m[:worksheet], range: m[:range])
      end

      def initialize(workbook:, worksheet:, range:)
        @workbook = workbook
        @worksheet = worksheet
        @range = range
      end

      # Whether this address names a range of cells. An address that stops at
      # a sheet or a book is a lookup, not an assignment target -- see the
      # spec on why `xl["Sheet1!"] = 0` filling a whole sheet is a hazard
      # rather than a convenience.
      def range?
        !@range.nil?
      end
    end
  end
end
