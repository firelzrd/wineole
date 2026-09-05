require_relative '../proxy'
require_relative '../errors'
require_relative 'passthrough'
require_relative 'vba_api'

module WineOLE
  module MSOffice
    # The tables and checks the three control collections share. Everything
    # here either maps a Ruby-side name onto the COM name Excel wants, or
    # refuses a call before Excel is touched -- the wrapper's job is to make
    # the passthrough's traps unreachable, not to add features to Excel.
    module Controls
      # Legacy form controls: kind => the Worksheet collection whose Add
      # makes one. Every collection takes (Left, Top, Width, Height).
      # EditBox is missing on purpose: it exists only on dialog sheets.
      FORM_KINDS = {
        button: 'Buttons',
        check_box: 'CheckBoxes',
        option_button: 'OptionButtons',
        list_box: 'ListBoxes',
        drop_down: 'DropDowns',
        spinner: 'Spinners',
        scroll_bar: 'ScrollBars',
        label: 'Labels',
        group_box: 'GroupBoxes'
      }.freeze

      # Shape.FormControlType => kind, for re-binding an existing shape.
      # 3 (xlEditBox) is absent for the reason above, so it maps to nil.
      FORM_CONTROL_TYPES = {
        0 => :button,
        1 => :check_box,
        2 => :drop_down,
        4 => :group_box,
        5 => :label,
        6 => :list_box,
        7 => :option_button,
        8 => :scroll_bar,
        9 => :spinner
      }.freeze

      # MSForms 2.0: kind => ProgID. Shorthand only -- any String is passed
      # to Excel verbatim as a ProgID, which is how a control outside this
      # table gets placed.
      MSFORMS_KINDS = {
        command_button: 'Forms.CommandButton.1',
        text_box: 'Forms.TextBox.1',
        combo_box: 'Forms.ComboBox.1',
        list_box: 'Forms.ListBox.1',
        check_box: 'Forms.CheckBox.1',
        option_button: 'Forms.OptionButton.1',
        toggle_button: 'Forms.ToggleButton.1',
        spin_button: 'Forms.SpinButton.1',
        scroll_bar: 'Forms.ScrollBar.1',
        label: 'Forms.Label.1',
        image: 'Forms.Image.1'
      }.freeze

      # A worksheet ActiveX control is two objects: the OLEObject Excel
      # wraps it in, and the MSForms control inside (OLEObject.Object).
      # These keys belong to the host; everything else goes inside.
      HOST_PROPS = %i[linked_cell list_fill_range visible print_object placement].freeze

      # A VBA identifier. Excel accepts any string as a shape name, but a
      # name that cannot appear in `Name_Click` is a control that can be
      # placed and never handled.
      VBA_NAME = /\A[A-Za-z][A-Za-z0-9_]{0,30}\z/.freeze

      def self.form_collection_for(kind)
        return FORM_KINDS[kind] if FORM_KINDS.key?(kind)

        if kind.is_a?(::String)
          raise ArgumentError,
            "form controls have no ProgID (#{kind.inspect} given): a String kind is for " \
            'sheet.activex and form.controls. Form control kinds are ' \
            "#{FORM_KINDS.keys.map(&:inspect).join(', ')}"
        end

        raise ArgumentError,
          "unknown form control kind #{kind.inspect} -- expected one of " \
          "#{FORM_KINDS.keys.map(&:inspect).join(', ')}. For an ActiveX control " \
          '(events reach Ruby) use sheet.activex'
      end

      def self.progid_for(kind)
        return kind if kind.is_a?(::String)
        return MSFORMS_KINDS[kind] if MSFORMS_KINDS.key?(kind)

        raise ArgumentError,
          "unknown ActiveX kind #{kind.inspect} -- expected one of " \
          "#{MSFORMS_KINDS.keys.map(&:inspect).join(', ')}, or pass a ProgID String " \
          'for any other registered control'
      end

      def self.kind_for_progid(progid)
        MSFORMS_KINDS.key(progid) || progid
      end

      def self.check_name!(name)
        return if name.is_a?(::String) && name.match?(VBA_NAME)

        raise ArgumentError,
          "name: must be a VBA identifier -- a letter, then letters, digits or " \
          "underscores, at most 31 characters -- because it becomes the `Name_Click` " \
          "handler name. Got #{name.inspect}"
      end

      def self.check_event!(event)
        return if event.is_a?(::String) && event.match?(VBA_NAME)

        raise ArgumentError,
          "an event name is a VBA identifier such as 'Click' or 'KeyDown'. Got #{event.inspect}"
      end

      # Excel allows two shapes with the same name in silence, and a UserForm
      # allows two controls with the same name in silence; after that
      # `Name_Click` means either. A lookup that raises is the free case.
      def self.check_free!(host, name, what)
        host.Item(name)
      rescue WineOLE::RemoteError
        nil
      else
        raise ArgumentError,
          "this #{what} already has a control named #{name.inspect}. Excel would add " \
          'a second one silently and then Name_Click would be ambiguous; pick another ' \
          'name or remove the existing control first'
      end

      # snake_case => PascalCase. COM matches names case-insensitively, so a
      # key already in PascalCase survives the round trip well enough.
      def self.pascal(key)
        key.to_s.split('_').map(&:capitalize).join
      end

      def self.put(target, key, value)
        target.public_send("#{pascal(key)}=", value)
      end

      # Either `at:` (a range on `sheet`, read for its box) or all four
      # points. Never a mix, never a partial box, never a default -- "wherever
      # Excel puts it" stays a passthrough behaviour. `sheet: nil` is a
      # UserForm, which has no cells for `at:` to name.
      def self.geometry(sheet:, at:, left:, top:, width:, height:)
        points = { left: left, top: top, width: width, height: height }
        given = points.reject { |_, v| v.nil? }.keys

        if !at.nil? && !given.empty?
          raise ArgumentError,
            "give either at: or left:/top:/width:/height:, not both (got at: #{at.inspect} " \
            "and #{given.inspect})"
        end

        unless at.nil?
          if sheet.nil?
            raise ArgumentError,
              'a UserForm has no cells; give left:, top:, width: and height: in points'
          end

          range = sheet[at].ole
          return [range.Left, range.Top, range.Width, range.Height]
        end

        if given.empty?
          raise ArgumentError,
            'no position given: pass at: "B2:C4" (a range on this sheet) or all four of ' \
            'left:, top:, width: and height: (points). There is no default position'
        end

        if given.length < 4
          raise ArgumentError,
            "left:, top:, width: and height: must all be given (missing #{(points.keys - given).inspect})"
        end

        [left, top, width, height]
      end
    end

    # One placed control, from any of the three families.
    #
    # TWO OBJECTS, ONE WRAPPER. A worksheet ActiveX control is an OLEObject
    # (Excel's host: Left, Top, Visible, LinkedCell) around an MSForms
    # control (Caption, Value, BackColor). `ole` is the host and
    # `ole_object` the MSForms control; unknown methods go to `ole_object`,
    # because that is where Caption and Value live for every family. Host
    # members are reached explicitly through `ctl.ole`. There is no lookup
    # order that tries both: `Left` exists on neither the inner object
    # (when hosted on a sheet) nor unambiguously on both. For the other two
    # families `ole` and `ole_object` are the same thing.
    #
    # Name note: `name` covers COM `Name` with the same value (M0). The
    # other lowercase methods here -- kind, family, ole, runtime, on, off,
    # events, vba -- are free on Button, OLEObject and the MSForms controls
    # (M0, measured against Excel 11). `ole_object` carries the `ole_`
    # prefix for the reason Proxy's meta-methods do: a bare `object` would
    # shadow COM `Object`, which on an MSForms control is a *different*
    # thing -- the raw control under the extender, with Caption but no Name
    # and no events (M0). `ctl.Object` still reaches that COM member.
    class Control
      include Passthrough

      attr_reader :name, :kind, :family, :ole_object

      # `vba:` is the writer a handler goes through: the BookVBA for a form
      # control (any standard module will do, so the wrapper's own), the
      # SheetVBA for worksheet ActiveX (Excel looks for `Name_Click` in the
      # sheet's own module and nowhere else), the BookVBA for a UserForm
      # control (`into:` the form module). `form:` is the Form a :userform
      # control belongs to; it owns the runtime instance.
      def initialize(name:, kind:, family:, ole:, ole_object:, vba:, form: nil)
        @name = name
        @kind = kind
        @family = family
        @ole = ole
        @ole_object = ole_object
        @writer = vba
        @form = form
      end

      # The Events object a block would be registered on; nil for a form
      # control, which has none.
      def events
        case @family
        when :activex  then @ole_object.ole_events
        when :userform then runtime.ole_events
        end
      end

      # A UserForm control's live counterpart on the form's default
      # instance -- the object that fires events and shows changes while the
      # form is loaded. The design-time control (`ole`) does neither.
      def runtime
        unless @family == :userform
          raise ArgumentError,
            "only a UserForm control has a runtime instance; this is a #{@family}"
        end

        @form.runtime_control(@name)
      end

      def on(event, args: true, &block)
        listenable.on(event, args: args, &block)
      end

      def off(name_or_subscription)
        listenable.off(name_or_subscription)
      end

      # Write one VBA handler. A form control fires only Click, so it takes
      # the body alone and is bound through OnAction; the other two take the
      # event name, and Excel finds the procedure by its `Name_Event` name in
      # the right module. `params:` is the parameter list, verbatim -- the
      # wrapper carries no signature table.
      #
      # The block is named `Name_Event`, so writing the same event again
      # replaces the handler (vba.write's own rule).
      def vba(event_or_body, body = nil, params: nil)
        if @family == :form_control
          unless body.nil?
            raise ArgumentError,
              'a form control fires only Click: call vba(body) with no event name'
          end

          macro = "#{@name}_Click"
          @writer.write("Sub #{macro}()\n#{indent(event_or_body)}\nEnd Sub", name: macro)
          @ole.OnAction = macro
        else
          if body.nil?
            raise ArgumentError,
              "vba(event, body) -- name the event, e.g. vba('Click', 'Range(\"A1\").Value = 1')"
          end

          Controls.check_event!(event_or_body)
          block = "#{@name}_#{event_or_body}"
          code = "Private Sub #{block}(#{params})\n#{indent(body)}\nEnd Sub"
          if @family == :activex
            @writer.write(code, name: block)
          else
            @writer.write(code, name: block, into: @form.name)
          end
        end
        self
      end

      private

      def passthrough_target
        @ole_object
      end

      def listenable
        if @family == :form_control
          raise ArgumentError,
            'form controls have no COM events; bind a macro with vba(...) or use ' \
            'sheet.activex for a control Ruby can listen to'
        end

        events
      end

      def indent(body)
        body.to_s.chomp.split(/\r?\n/, -1).map { |line| line.empty? ? line : "    #{line}" }.join("\n")
      end
    end

    # `sheet.form_controls`: the Forms-toolbar controls. Cheap to place (a
    # `Buttons.Add` plus `OnAction` measured 3 ms) and they save with the
    # workbook, but they raise no COM events -- a handler is a macro named
    # by OnAction, so `vba(body)` is the only way to react to one.
    #
    # Name note (M0): `form_controls` is free on Worksheet.
    class FormControls
      def initialize(sheet)
        @sheet = sheet
      end

      # Order matters and each step exists to make a specific passthrough
      # trap unreachable: kind and name are checked before geometry so a
      # typo fails without a round trip; the free-name check runs against
      # Shapes, which sees every family; the rename and the properties run
      # after Add with the new control deleted if either fails, so a refused
      # property does not leave an unnamed button behind.
      def add(kind, name:, at: nil, left: nil, top: nil, width: nil, height: nil, **props)
        collection = Controls.form_collection_for(kind)
        Controls.check_name!(name)
        l, t, w, h = Controls.geometry(sheet: @sheet, at: at, left: left, top: top, width: width, height: height)
        Controls.check_free!(@sheet.ole.Shapes, name, 'sheet')

        ole = @sheet.ole.public_send(collection).Add(l, t, w, h)
        begin
          ole.Name = name
          props.each { |key, value| Controls.put(ole, key, value) }
        rescue StandardError
          ole.Delete
          raise
        end
        Control.new(name: name, kind: kind, family: :form_control, ole: ole, ole_object: ole, vba: book_vba)
      end

      # Re-bind a control that is already on the sheet. nil when there is no
      # shape of that name, or the shape is not a form control (an ActiveX
      # shape raises on FormControlType; an EditBox maps to nil).
      def [](name)
        kind = Controls::FORM_CONTROL_TYPES[@sheet.ole.Shapes.Item(name).FormControlType]
        return nil if kind.nil?

        ole = @sheet.ole.public_send(Controls::FORM_KINDS[kind]).Item(name)
        Control.new(name: name, kind: kind, family: :form_control, ole: ole, ole_object: ole, vba: book_vba)
      rescue WineOLE::RemoteError
        nil
      end

      private

      # A form control's macro can live in any standard module, so it goes
      # in the wrapper's own module of the parent workbook. Only `write` is
      # used, and paths never are, so convert_paths is moot.
      def book_vba
        @book_vba ||= BookVBA.new(@sheet.ole.Parent, convert_paths: false)
      end
    end

    # `sheet.activex`: OLEObjects hosting an MSForms control (or any other
    # registered control, by ProgID). Each is two COM objects -- see
    # Control's class comment -- and the properties given at placement are
    # routed accordingly: HOST_PROPS to the OLEObject, the rest inside.
    #
    # The five named arguments to OLEObjects.Add are not optional on Excel
    # 11: with only Left and Top it fails with 0x800A03EC. Geometry has
    # already guaranteed all four points by the time Add is called.
    #
    # Name note (M0): `activex` is free on Worksheet.
    class ActiveXControls
      def initialize(sheet)
        @sheet = sheet
      end

      def add(kind, name:, at: nil, left: nil, top: nil, width: nil, height: nil, **props)
        progid = Controls.progid_for(kind)
        Controls.check_name!(name)
        l, t, w, h = Controls.geometry(sheet: @sheet, at: at, left: left, top: top, width: width, height: height)
        Controls.check_free!(@sheet.ole.Shapes, name, 'sheet')

        ole = @sheet.ole.OLEObjects.Add(ClassType: progid, Left: l, Top: t, Width: w, Height: h)
        begin
          ole.Name = name
          inner = ole.Object
          props.each do |key, value|
            Controls.put(Controls::HOST_PROPS.include?(key) ? ole : inner, key, value)
          end
        rescue StandardError
          ole.Delete
          raise
        end
        Control.new(name: name, kind: kind, family: :activex, ole: ole, ole_object: inner, vba: @sheet.vba)
      end

      def [](name)
        ole = @sheet.ole.OLEObjects.Item(name)
        Control.new(name: name, kind: Controls.kind_for_progid(ole.progID), family: :activex,
                    ole: ole, ole_object: ole.Object, vba: @sheet.vba)
      rescue WineOLE::RemoteError
        nil
      end
    end

    # `form.controls`: MSForms controls on a UserForm, placed on the
    # design-time Designer. Points only -- a form has no cells for `at:`.
    # MSForms takes the name at Add time, so there is no rename step, and
    # the box is four property puts after it (Controls.Add has no position
    # arguments).
    #
    # Name note (M0): `controls` shadows `Designer.Controls` deliberately;
    # the COM collection is still reachable as `form.ole.Controls`.
    class UserFormControls
      def initialize(form, book_vba)
        @form = form
        @book_vba = book_vba
      end

      def add(kind, name:, at: nil, left: nil, top: nil, width: nil, height: nil, **props)
        progid = Controls.progid_for(kind)
        Controls.check_name!(name)
        l, t, w, h = Controls.geometry(sheet: nil, at: at, left: left, top: top, width: width, height: height)
        Controls.check_free!(@form.ole.Controls, name, 'UserForm')

        ole = @form.ole.Controls.Add(progid, name)
        begin
          ole.Left = l
          ole.Top = t
          ole.Width = w
          ole.Height = h
          props.each { |key, value| Controls.put(ole, key, value) }
        rescue StandardError
          @form.ole.Controls.Remove(name)
          raise
        end
        Control.new(name: name, kind: kind, family: :userform, ole: ole, ole_object: ole, vba: @book_vba, form: @form)
      end

      # kind is nil: a placed MSForms control does not report the ProgID it
      # was made from.
      def [](name)
        ole = @form.ole.Controls.Item(name)
        Control.new(name: name, kind: nil, family: :userform, ole: ole, ole_object: ole, vba: @book_vba, form: @form)
      rescue WineOLE::RemoteError
        nil
      end
    end
  end
end
