require_relative '../../wineole'
require_relative '../proxy'
require_relative 'passthrough'
require_relative 'address'
require_relative 'range'
require_relative 'sheet'
require_relative 'book'

module WineOLE
  module MSOffice
    # Wraps a COM Excel.Application: the lifecycle (.create / .connect /
    # .connect_or_create / .run) and the entry point of the addressing DSL
    # that Address / Sheet / Book / Range build on.
    #
    # `version` is deliberately not defined -- see the comment on @version
    # in #initialize.
    #
    # `run` here is a *class* method (Excel.run), never an instance one:
    # COM's own Application.Run (the macro runner) is real -- measured on a
    # live Excel 11, `xl.run` answers DISP_E_EXCEPTION, meaning COM tried
    # to actually invoke it rather than reporting an unknown name. A class
    # method lives in a completely separate namespace from an instance
    # method of the same spelling, so `Excel.run` shadows nothing; adding
    # an *instance* `run` would.
    class Excel
      include Passthrough

      APPLICATION = 'Excel.Application'.freeze

      # What the bridge runs on the way out of an auto-created instance,
      # once its last user releases the root: suppress prompts, then quit.
      # Declared here, once, and handed to the bridge at construction time
      # (via Client#create/connect/connect_or_create's `cleanup:` kwarg)
      # rather than run from Ruby -- the bridge is the only party that knows
      # when the LAST user of a shared root has let go of it, so it is the
      # only party that can decide whether these steps should run at all.
      CLEANUP_STEPS = { steps: [['DisplayAlerts=', false], ['Quit']] }.freeze

      def self.create(client: WineOLE.default_client, convert_paths: true)
        new(client.create(APPLICATION, cleanup: CLEANUP_STEPS), client: client, convert_paths: convert_paths)
      end

      # Declares the same steps .create does. That is correct, not an
      # oversight: the steps are a property of this instance, but they only
      # ever RUN when the bridge's record is auto_created, which is false
      # here unless connect_or_create's own fallback is what created it. A
      # human's Excel reached via .connect is never auto_created, so Quit
      # never fires for it -- this class used to enforce that with an
      # `ole_created?` check of its own; the bridge enforces it now, one
      # layer down.
      def self.connect(client: WineOLE.default_client, convert_paths: true)
        new(client.connect(APPLICATION, cleanup: CLEANUP_STEPS), client: client, convert_paths: convert_paths)
      end

      def self.connect_or_create(client: WineOLE.default_client, convert_paths: true)
        new(client.connect_or_create(APPLICATION, cleanup: CLEANUP_STEPS), client: client, convert_paths: convert_paths)
      end

      # Runs the block with an Excel application, then releases it.
      #
      # Deciding whether that release actually quits Excel is no longer this
      # method's job. create/connect/connect_or_create declare CLEANUP_STEPS
      # at construction time, and the bridge is the one that knows when the
      # LAST user of an auto-created root lets go of it -- that is when it
      # runs those steps. Attaching to an Excel somebody already had open
      # (.connect) is never auto-created, so those steps never fire for it;
      # quitting it out from under them would throw away their unsaved work.
      def self.run(mode = :connect_or_create, **options)
        xl = case mode
             when :create             then create(**options)
             when :connect            then connect(**options)
             when :connect_or_create  then connect_or_create(**options)
             else raise ArgumentError, "unknown mode #{mode.inspect}"
             end
        begin
          yield xl
        ensure
          xl.ole_release
        end
      end

      def initialize(proxy, client:, convert_paths: true)
        @ole = proxy
        @client = client
        @convert_paths = convert_paths
        # COM resolves names case-insensitively, so `xl.version` already
        # reaches COM's own Version and returns e.g. "11.0" -- measured.
        # Defining a `version` method here would only shadow that with
        # something that does the same thing, so this class deliberately
        # does not have one. Captured once here, rather than read fresh
        # per lookup, so every Sheet/Book this object builds sees one
        # stable answer without a round trip each time.
        @version = @ole.Version
      end

      # Release the underlying Application. On the last user of an auto-created
      # instance this is what quits Excel (the bridge runs DisplayAlerts=false
      # then Quit); on a connected instance it simply detaches. Public because
      # `run` is not the only way to get one of these, and a caller managing
      # the lifecycle by hand needs the same call available.
      def ole_release
        @ole.ole_release
      end

      # Keep this Excel running after this program leaves -- e.g. a report left
      # on screen for a human. Revokes the bridge's permission to quit it.
      def leave_open
        @ole.ole_leave_open
      end

      def [](*keys)
        lookup(keys)
      end

      # Resolves the same way #[] does; raises unless that resolves all the
      # way to a Range. `xl["Sheet1!"] = 0` silently filling a whole sheet
      # is exactly the hazard this guards against -- a typo that stops at a
      # sheet or a book is otherwise indistinguishable from that "fill
      # everything" request (Spec Sec 4.5).
      def []=(*args)
        value = args.pop
        target = lookup(args)
        unless target.is_a?(Range)
          raise ArgumentError,
            "#{args.inspect} does not resolve to a range -- assignment needs an address " \
            'with an explicit range; a bare sheet or workbook lookup is get-only'
        end
        target.write(value)
      end

      # COM's Application has no Show/Hide (measured: both answer
      # DISP_E_UNKNOWNNAME), so defining these does not shadow a working
      # COM call the way a single-word name with a real COM counterpart
      # would (Sec 4.6).
      def show
        @ole.Visible = true
      end

      def hide
        @ole.Visible = false
      end

      # Suppresses Excel's save/overwrite/etc. modal prompts for the
      # duration of the block, then restores whatever DisplayAlerts was
      # set to beforehand -- not a hardcoded true. msoffice.rb's own
      # version restores a hardcoded true in its ensure, which would
      # silently turn a caller's own `DisplayAlerts = false` back on the
      # moment this method returns; reading the value first instead means
      # an outer caller who had already turned it off keeps it off.
      #
      # The underscore makes the name safe regardless of whether COM has a
      # same-named member (Sec 4.6): `no_alert` would not resolve to
      # anything on Application even as a single word, since COM does not
      # strip underscores when matching names.
      def no_alert
        raise ArgumentError, 'no_alert needs a block -- the flag is restored when it returns' unless block_given?

        previous = @ole.DisplayAlerts
        begin
          @ole.DisplayAlerts = false
          yield
        ensure
          restore(:DisplayAlerts=, previous)
        end
      end

      # Same restore-what-was-there discipline as #no_alert, for screen
      # redraws instead of alert dialogs.
      def no_update
        raise ArgumentError, 'no_update needs a block -- the flag is restored when it returns' unless block_given?

        previous = @ole.ScreenUpdating
        begin
          @ole.ScreenUpdating = false
          yield
        ensure
          restore(:ScreenUpdating=, previous)
        end
      end

      private

      # Put a flag back, and never let failing to do so become the caller's
      # problem.
      #
      # An exception raised inside `ensure` REPLACES whatever the block
      # raised. A failed restore would therefore destroy the caller's own
      # error and report the cleanup instead.
      #
      # And the ordinary reason a restore fails is that the block ended the
      # very thing being restored: an Application that quits or disconnects
      # mid-block (whether from a caller's own `no_alert { xl.Quit }`, or
      # from underneath this process entirely) leaves nothing for `no_alert`'s
      # ensure to put DisplayAlerts back on. There is no state left to
      # restore, and nothing worth telling anyone.
      #
      # Swallowed unconditionally rather than only for a "the object is
      # gone" error: over this bridge that arrives as a RemoteError wrapping
      # an HRESULT string (0x800706BE is a normal transient right after
      # Quit), so telling the two apart means matching on those strings --
      # brittle, and beside the point, since raising out of `ensure` is
      # wrong even when the failure really is transient.
      def restore(setter, value)
        @ole.public_send(setter, value)
      rescue StandardError
        nil
      end

      def lookup(keys)
        return cell_range(*keys) if keys.length == 2 && keys.all? { |k| k.is_a?(::Integer) }

        unless keys.length == 1 && keys.first.is_a?(::String)
          raise ArgumentError, "unsupported index #{keys.inspect}"
        end

        str = keys.first
        addr = Address.parse(str, @version)

        # Not an address at all -- msoffice.rb's own fallback: treat the
        # whole string as a raw worksheet name. This is exactly why
        # Address.parse returns nil instead of raising.
        return Sheet.new(@ole.Worksheets.Item(str), version: @version) if addr.nil?

        resolve(addr, str)
      end

      # xl[row, col]: Cells on the active sheet. Goes through #resolve's
      # same no-Select path (via the Sheet it builds), one Application
      # round trip for ActiveSheet plus whatever Sheet#[] costs.
      def cell_range(row, col)
        Sheet.new(active_worksheet(@ole, "xl[#{row}, #{col}]"), version: @version)[row, col]
      end

      # Resolves a parsed Address against this Application.
      #
      # Deliberately never calls Select. msoffice.rb has to, because it
      # reaches a range through Application.Range, which addresses
      # whatever sheet happens to be active. Resolving against the
      # worksheet object itself needs no active state, so a read here does
      # not mutate the caller's selection as a side effect, and it costs
      # one fewer round trip. Measured against a live Excel 11:
      # `book.Worksheets.Item(1).Range('A1').Value` left ActiveSheet
      # exactly where it was, both before and after.
      def resolve(addr, str)
        workbook_ole = addr.workbook.nil? ? nil : resolve_workbook(addr.workbook, str)

        worksheet_part = addr.worksheet
        # A bare range ("A1:B2", no "!") names no worksheet at all -- it
        # still needs one to be looked up against, so it implicitly means
        # the active sheet, the same way xl[row, col] does.
        worksheet_part = '' if worksheet_part.nil? && addr.range?

        if worksheet_part
          worksheet_ole = resolve_worksheet(worksheet_part, workbook_ole || @ole, str)
          sheet = Sheet.new(worksheet_ole, version: @version)
          return addr.range? ? sheet[addr.range] : sheet
        end

        # The grammar only leaves worksheet_part nil when addr is the bare
        # "[workbook]" form -- Address#worksheet is nil and Address#range?
        # is false together only there -- so workbook_ole is always set by
        # this point (Sec 4.5, get-only for a lookup with no range).
        Book.new(workbook_ole, client: @client, version: @version, convert_paths: @convert_paths)
      end

      def resolve_workbook(part, str)
        if part.empty?
          active_workbook(str)
        elsif part.casecmp?(':new')
          @ole.Workbooks.Add
        else
          @ole.Workbooks.Item(part)
        end
      end

      def resolve_worksheet(part, container, str)
        if part.empty?
          active_worksheet(container, str)
        elsif part.casecmp?(':new')
          worksheets = container.Worksheets
          worksheets.Add(After: worksheets.Item(worksheets.Count))
        elsif part.casecmp?(':first')
          container.Worksheets.Item(1)
        elsif part.casecmp?(':last')
          worksheets = container.Worksheets
          worksheets.Item(worksheets.Count)
        elsif part.match?(/\A\d+\z/)
          container.Worksheets.Item(part.to_i)
        else
          container.Worksheets.Item(part)
        end
      end

      # ActiveWorkbook/ActiveSheet are nil on a fresh Excel with nothing
      # open yet (measured: Workbooks.Count == 0 makes both nil). Without
      # this check that turns into a NoMethodError raised from deep inside
      # this class instead of something a caller can act on.
      #
      # RuntimeError rather than ArgumentError: the address string itself
      # was fine -- it is the application's current state (no open
      # workbook/sheet) that cannot satisfy it.
      def active_workbook(str)
        @ole.ActiveWorkbook or raise RuntimeError, "no active workbook -- #{str.inspect} needs one open"
      end

      def active_worksheet(container, str)
        container.ActiveSheet or raise RuntimeError, "no active worksheet -- #{str.inspect} needs one open"
      end
    end
  end
end
