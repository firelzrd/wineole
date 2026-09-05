require 'tmpdir'
require_relative 'paths'
require_relative 'vba'
require_relative 'vba_block'

module WineOLE
  module MSOffice
    # The VBA surface of a workbook, reached as `book.vba`.
    #
    # TWO GRANULARITIES LIVE HERE and the names are what keep them apart.
    # `write` and `remove` act on a named BLOCK inside a module -- the span
    # this wrapper owns, delimited by sentinel comments, which may sit in a
    # module full of code somebody else wrote. `add_component`,
    # `remove_component`, `import` and `export` act on whole COMPONENTS.
    #
    # The verbs are not interchangeable either, and the difference is
    # deliberate. `write` is an upsert: writing the same name again replaces
    # that block, so it is idempotent. `add_component` is create-only and
    # refuses a name that is already taken -- "overwriting" a component
    # would destroy whatever a person put in it, so there is no way to ask
    # for that.
    #
    # WHERE CODE GOES DECIDES WHETHER IT CAN BE CALLED. Measured against a
    # live Excel 11:
    #
    #   standard module   Run("Name") works and returns the value;
    #                     a worksheet formula =Name() works too.
    #                     Private does NOT hide it from either.
    #   ThisWorkbook,     Run("Name") fails ("macro not found"). Run with
    #   a worksheet,      the module qualified -- Run("Sheet1.Name") --
    #   a UserForm        runs it but hands back nil, so a Function's
    #                     return value cannot be collected. =Name() is
    #                     #NAME?.
    #
    # So code meant to be called belongs in a standard module (which is
    # where `into: nil` puts it), and code in a sheet or ThisWorkbook module
    # is there to be reached by Excel itself -- an ActiveX control's
    # `_Click`, a workbook event.
    class BookVBA
      # The module this wrapper makes for itself when `into:` is not given.
      DEFAULT_MODULE = 'WineOLE'.freeze

      # VBComponents.Add's type argument. 100 (a Document module -- a
      # worksheet or ThisWorkbook) is deliberately absent: Excel owns those
      # and neither creates nor destroys them on request.
      KINDS = { standard: 1, class: 2, form: 3 }.freeze

      # The Type of a component Excel owns. Cannot be added, cannot be
      # removed -- only emptied.
      DOCUMENT_TYPE = 100

      def initialize(ole, convert_paths:)
        @ole = ole
        @convert_paths = convert_paths
      end

      # Put a named block of VBA into this workbook.
      #
      # `into:` names an EXISTING component -- a UserForm, ThisWorkbook, a
      # worksheet's module, a module made with #add_component. It is not
      # created on demand: a typo would otherwise become a new module in
      # silence. Without it the block goes in this wrapper's own module,
      # which is created on demand because its name is not the caller's to
      # get wrong.
      #
      # The block is what this wrapper owns -- writing the same name again
      # replaces it, and nothing else in the module is touched. That matters
      # because the module may hold code somebody wrote by hand.
      def write(code, name: 'main', into: nil)
        VBABlock.write(target_module(into), name, code)
        self
      end

      # Remove a named block. When `from:` is given, only the block is
      # removed -- a component the caller named is the caller's, never swept
      # up here (a UserForm can be deleted, unlike ThisWorkbook, so this has
      # to be a rule rather than an accident of what COM allows). Use
      # #remove_component to delete one deliberately.
      #
      # Without `from:`, this wrapper's own module is the target, and when
      # it has nothing but whitespace left the module goes too -- an empty
      # module is litter, and that one IS ours to clean up.
      #
      # VBABlock.remove already fetched the whole body to find the block; it
      # hands back the remaining lines so this does not have to fetch the
      # body a second time just to ask whether it is empty.
      def remove(name, from: nil)
        if from
          VBABlock.remove(named_component!(from).CodeModule, name)
          return self
        end

        component = existing_component(DEFAULT_MODULE)
        return self if component.nil?

        remaining = VBABlock.remove(component.CodeModule, name)
        project.VBComponents.Remove(component) if remaining && VBABlock.blank_lines?(remaining)
        self
      end

      # Create an empty component. Create-only on purpose: see the class
      # comment on why there is no "overwrite a component".
      #
      # The existence check has to come BEFORE Add, not after, and that is
      # measured rather than defensive: Add and the rename that follows it
      # are not atomic. Adding under a taken name SUCCEEDS, and only the
      # rename fails (0x80020009), leaving a stray `Module1` behind that
      # nobody asked for. Checking first is what keeps the failure clean.
      def add_component(name, kind: :standard)
        type = KINDS[kind]
        unless type
          raise ArgumentError,
            "unknown component kind #{kind.inspect} -- expected one of #{KINDS.keys.inspect}. " \
            'A worksheet module and ThisWorkbook are not on that list because Excel ' \
            'owns them; they exist already and cannot be made'
        end

        if existing_component(name)
          raise ArgumentError,
            "this workbook already has a VBA component named #{name.inspect}. " \
            'add_component never overwrites one -- that would destroy whatever is ' \
            'in it. Remove it first, or use write(into:) to put a block inside it'
        end

        component = project.VBComponents.Add(type)
        begin
          component.Name = name
        rescue WineOLE::RemoteError
          # Add succeeded and the rename did not, so the component exists
          # under a name nobody chose. Take it back out rather than leaving
          # the litter the pre-check exists to prevent.
          project.VBComponents.Remove(component)
          raise ArgumentError,
            "Excel refused #{name.inspect} as a component name. VBA names start with a " \
            'letter and hold letters, digits and underscores, up to 31 characters'
        end
        component
      end

      # Delete a component outright, with whatever is inside it.
      def remove_component(name)
        component = named_component!(name)
        if component.Type == DOCUMENT_TYPE
          raise ArgumentError,
            "#{name.inspect} is a module Excel owns (a worksheet's, or ThisWorkbook's) " \
            'and cannot be deleted -- it exists for as long as the sheet or the workbook ' \
            'does. To take this wrapper\'s code back out of it, use remove(name, from:)'
        end

        project.VBComponents.Remove(component)
        self
      end

      # Read a VBA source file on this machine into the project as a new
      # component. Excel's own Import does the work, so the component's name
      # and kind come from the file.
      #
      # `encoding:` skips the detection when the caller already knows; the
      # rules the default follows are on VBA.detect_encoding.
      def import(path, encoding: nil)
        local_bridge!('import')
        reject_dotdot!(path)
        source = encoding || VBA.detect_encoding(path)
        text = decode(::File.binread(path), source, path, guessed: encoding.nil?)

        Dir.mktmpdir('wineole-vba') do |dir|
          staged = ::File.join(dir, ::File.basename(path))
          ::File.binwrite(staged, to_codepage(text, path))
          project.VBComponents.Import(Paths.to_wine(staged))
        end
        self
      end

      # Write a component out as a file on this machine, in UTF-8 with LF --
      # what Excel produces is the ANSI codepage with CRLF, and the
      # destination is a Linux path.
      def export(name, path)
        local_bridge!('export')
        reject_dotdot!(path)
        component = named_component!(name)
        Dir.mktmpdir('wineole-vba') do |dir|
          staged = ::File.join(dir, ::File.basename(path))
          component.Export(Paths.to_wine(staged))
          text = ::File.binread(staged).force_encoding(VBA.codepage).encode('UTF-8')
          ::File.write(path, text.gsub("\r\n", "\n"))
        end
        self
      end

      # The workbook's VBA project, or an error that says what to do about
      # it. The HRESULT and the message are both useless for telling this
      # condition apart -- 0x800A03EC is what a rejected NumberFormat gives
      # too, and the text is localized -- so the registry is what turns a
      # refusal into advice.
      def project
        @ole.VBProject
      rescue WineOLE::RemoteError
        VBA.denied!
      end

      private

      # The other direction, and it had been left bare: a file that reads
      # cleanly can still hold a character the codepage cannot store, and
      # #encode raised Encoding::UndefinedConversionError naming neither the
      # file nor the way out. Same rule and same words as the string path --
      # refuse rather than let Excel substitute in silence.
      def to_codepage(text, path)
        text.encode(VBA.codepage)
      rescue ::Encoding::UndefinedConversionError => e
        VBA.unrepresentable!(e.error_char, path.to_s)
      end

      # Both encodings can be wrong at once: the bytes are not valid UTF-8
      # (which is what put us on the codepage branch) and not valid in the
      # codepage either. Nothing can be inferred from that, so say so --
      # a bare Encoding::InvalidByteSequenceError names neither the file
      # nor why that encoding was the one tried.
      def decode(raw, source, path, guessed:)
        raw.dup.force_encoding(source).encode('UTF-8').sub(/\A\uFEFF/, '')
      rescue ::Encoding::InvalidByteSequenceError, ::Encoding::UndefinedConversionError => e
        why =
          if guessed
            "#{source} was tried because the file has no BOM and its bytes are not valid " \
            "UTF-8, which rules UTF-8 out -- but they are not valid #{source} either, so " \
            'there is nothing left to infer from. Pass `encoding:` if you know what this is'
          else
            "you passed encoding: #{source.inspect}, and the file's bytes are not valid in it"
          end
        raise ArgumentError, "cannot read #{path} as #{source} (#{e.message}). #{why}"
      end

      # These hand Excel a path to a file on this machine. When the bridge
      # is somewhere else that path means that machine's filesystem, and
      # there is no sensible thing to do with it.
      #
      # Keying this off @convert_paths (rather than a separate "is this
      # loopback" flag) is deliberate, and it means a caller who passed
      # convert_paths: false for a legitimate reason on a *loopback* bridge
      # gets refused here too, with a message that says "needs the bridge to
      # be on this machine" when it already is. That is the wrong reason but
      # never the wrong direction: this can only over-refuse a bridge that
      # would have worked, never under-refuse one that would not.
      def local_bridge!(what)
        return if @convert_paths

        raise ArgumentError,
          "#{what} needs the bridge to be on this machine: it stages a file " \
          'and hands Excel the path, which means nothing on another host'
      end

      # File.basename of a path ending in ".." is literally "..". For import
      # that path is read directly, so it already raises Errno::EISDIR of its
      # own accord -- but a bare EISDIR does not say why. For export the same
      # basename feeds File.join(dir, "..") when staging, which resolves to
      # dir's *parent* rather than anywhere under our own tmpdir, and Export
      # ends up trying to write a file over that directory -- also EISDIR,
      # also confusing. Nothing escapes and nothing is corrupted either way;
      # this just gives both a clear error instead of a bare Errno.
      def reject_dotdot!(path)
        return unless ::File.basename(path) == '..'

        raise ArgumentError, "#{path.inspect} is not a usable file path (its basename is \"..\")"
      end

      def existing_component(name)
        project.VBComponents.Item(name)
      rescue WineOLE::RemoteError
        nil
      end

      def target_module(into)
        return own_module.CodeModule if into.nil?

        named_component!(into).CodeModule
      end

      # This wrapper's own module, made on demand. Unlike a name the caller
      # passed, this one cannot be a typo, so creating it silently is safe.
      def own_module
        found = existing_component(DEFAULT_MODULE)
        return found unless found.nil?

        component = project.VBComponents.Add(KINDS[:standard])
        component.Name = DEFAULT_MODULE
        component
      end

      def named_component!(name)
        found = existing_component(name)
        return found unless found.nil?

        raise ArgumentError,
          "this workbook has no VBA component named #{name.inspect}. " \
          'Components are UserForms, ThisWorkbook, worksheet modules and ' \
          'standard modules; add one with add_component, or omit `into:` ' \
          'to use this wrapper\'s own module'
      end
    end

    # The VBA surface of one worksheet, reached as `sheet.vba`.
    #
    # Blocks only. A worksheet's module cannot be created or deleted -- Excel
    # makes it with the sheet and destroys it with the sheet -- so the
    # component methods are absent here rather than present and always
    # failing. Removing the last block empties the module; it does not
    # remove it.
    class SheetVBA
      def initialize(ole)
        @ole = ole
      end

      def write(code, name: 'main')
        VBABlock.write(code_module, name, code)
        self
      end

      def remove(name)
        VBABlock.remove(code_module, name)
        self
      end

      private

      # A worksheet's handlers live in the worksheet's own code module --
      # that is where Excel looks for `<ActiveX control>_Click`. The module
      # is named by the sheet's CodeName, inside the parent workbook's
      # project.
      #
      # THE ORDER OF THESE TWO LINES IS LOAD-BEARING, and it is not obvious.
      # Worksheet.CodeName comes back as "" until something has touched that
      # workbook's VBProject -- measured: "" before, "Sheet3" after, and
      # "Sheet3" on every read from then on. Reading the name first and then
      # the project (which looks like the same code, and is what extracting
      # a local naturally produces) hands VBComponents.Item("") and gets
      # 0x800A0009, "index out of range". So the project is fetched first,
      # on purpose, and the name after it.
      #
      # The empty check is what keeps that from becoming a silent trap
      # again: if this ever stops holding, it fails saying why instead of
      # failing as a bare COM index error.
      def code_module
        vb_project = project
        code_name = @ole.CodeName.to_s
        if code_name.empty?
          raise VBA::Error,
            'this worksheet reports no CodeName even after its VBProject was opened, ' \
            'so there is no way to find its code module'
        end

        vb_project.VBComponents.Item(code_name).CodeModule
      end

      # Only the VBProject fetch is the denial. Wrapping the lookup that
      # follows it in the same rescue would report any other COM failure --
      # a component that is not there, a module that will not open -- as
      # "turn on AccessVBOM", which is advice for a condition the caller is
      # not in.
      def project
        @ole.Parent.VBProject
      rescue WineOLE::RemoteError
        VBA.denied!
      end
    end
  end
end
