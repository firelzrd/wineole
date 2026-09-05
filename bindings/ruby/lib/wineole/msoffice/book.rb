require 'tmpdir'
require_relative '../proxy'
require_relative 'passthrough'
require_relative 'paths'
require_relative 'sheet'
require_relative 'vba_api'
require_relative 'forms'

module WineOLE
  module MSOffice
    # Wraps a COM Workbook.
    #
    # Name note (measured against a live Excel 11): `sheet`, `each_sheet`,
    # `save_as` and `local_path` are all free -- COM answers
    # DISP_E_UNKNOWNNAME for each. `close` is not: COM resolves member names
    # case-insensitively, and `Workbook.Close` already exists. It is a
    # deliberate shadow anyway (see #close below). `local_file` is free by
    # the same construction as `save_as`: COM does not strip underscores
    # when matching names, so it never collides with `Workbook.FullName`.
    class Book
      include Passthrough

      # `convert_paths` is the caller's opt-out; whether converting is even
      # meaningful is a separate question (a remote bridge's own filesystem
      # means nothing to this machine). Combining them once here means a
      # caller cannot talk a remote bridge into converting by passing
      # `convert_paths: true` -- Paths.convertible? still says no.
      def initialize(proxy, client:, version:, convert_paths: true)
        @ole = proxy
        @version = version
        @convert_paths = convert_paths && Paths.convertible?(client: client)
      end

      # Worksheets, not Sheets: Sheets also includes chart sheets, which
      # this wrapper's Sheet class does not model.
      def sheet(name_or_index)
        Sheet.new(@ole.Worksheets.Item(name_or_index), version: @version)
      end

      def each_sheet
        return to_enum(:each_sheet) unless block_given?

        worksheets = @ole.Worksheets
        (1..worksheets.Count).each do |i|
          yield Sheet.new(worksheets.Item(i), version: @version)
        end
      end

      # Takes only the path. A caller needing FileFormat and the rest of
      # COM's SaveAs arguments uses the passthrough `book.SaveAs(...)`.
      def save_as(path)
        target = @convert_paths ? Paths.to_wine(path) : path
        @ole.SaveAs(target)
      end

      # COM's Workbook.Path is the *containing folder*, not the file --
      # local_path deliberately names the same thing this wrapper's way, in
      # Linux form. The file's own path is #local_file.
      #
      # An unsaved book's Path is "" -- Paths.to_local returns that
      # unchanged without shelling out to winepath, so this never runs it
      # for nothing.
      def local_path
        @convert_paths ? Paths.to_local(@ole.Path) : @ole.Path
      end

      # The file's own path, in Linux form -- what local_path is not.
      # Gated by the same @convert_paths (loopback-only) rule as
      # local_path; calling Paths.to_local(book.FullName) directly instead
      # skips that gate and runs a local winepath over what may be a
      # *remote* bridge's Wine path, silently producing a path that refers
      # to a filesystem this machine does not have (Spec Sec 4.7: a wrong
      # conversion that happens silently is worse than one that visibly
      # does not happen).
      #
      # Measured against a live Excel 11: FullName is the file
      # (`Z:\tmp\wineole_item_probe.xls` where Path is the folder,
      # `Z:\tmp`); an unsaved book's FullName is the bare in-memory name
      # ("Book1"), matching its Path of "".
      def local_file
        @convert_paths ? Paths.to_local(@ole.FullName) : @ole.FullName
      end

      # A deliberate shadow of COM's Workbook.Close. Close with no
      # SaveChanges argument can raise a modal save-changes prompt, which
      # under Wine is a hang; close(save: false) turns that hazard into an
      # explicit parameter instead. The raw member stays reachable as
      # `book.Close(...)` (exact PascalCase) and as `book.ole.Close(...)`.
      def close(save: false)
        @ole.Close(save)
      end

      # This workbook's VBA surface: blocks, components, import and export.
      # `book.vba.write(code, name: 'helpers')` is the common call; see
      # BookVBA for the block-vs-component split and for where code has to
      # live to be callable at all.
      #
      # Name note: `vba` is a bare lowercase word, so it would shadow a COM
      # member spelled the same. Workbook has none -- the member it has is
      # `VBProject`, still reachable as `book.vba.project`.
      def vba
        @vba ||= BookVBA.new(@ole, convert_paths: @convert_paths)
      end

      # This workbook's UserForms: `forms.add('AppForm')`, `forms['AppForm']`.
      #
      # Name note (M0, measured against Excel 11): `forms` is free on
      # Workbook.
      def forms
        @forms ||= Forms.new(self)
      end
    end
  end
end
