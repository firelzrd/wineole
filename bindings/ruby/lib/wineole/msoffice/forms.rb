require_relative '../proxy'
require_relative '../errors'
require_relative 'passthrough'
require_relative 'vba_api'
require_relative 'controls'

module WineOLE
  module MSOffice
    # `book.forms`: the UserForms of a workbook. A UserForm is a VBComponent
    # of type 3 with a Designer; the form itself exists only while VBA has
    # it loaded, as its default instance. The wrapper reaches that instance
    # through four generated procedures in its own module, because a
    # UserForm's default instance has no COM name a client can ask for.
    #
    # Name note (M0): `forms` is free on Workbook.
    class Forms
      USERFORM_TYPE = 3

      def initialize(book)
        @book = book
      end

      # The name is checked here as well as in add_component, because it
      # also becomes part of four procedure names.
      def add(name)
        Controls.check_name!(name)
        component = @book.vba.add_component(name, kind: :form)
        begin
          Form.new(name, component, @book)
        rescue StandardError
          # Form.new's helper write failed after the component was already
          # created; take the component back out rather than leaving it
          # behind with no helper to show or unload it.
          @book.vba.remove_component(name)
          raise
        end
      end

      # Re-bind an existing UserForm (one made earlier, or one in a workbook
      # that was opened). nil for a missing name or a component that is not
      # a UserForm. AccessDenied from `project` passes through.
      #
      # Only the lookup is rescued, deliberately: Form.new writes the helper
      # block, and that write can be refused (a locked project, a refused
      # AddFromString). Reporting a real failure as "not found" would send
      # the caller looking for a form that is right there.
      def [](name)
        component = begin
          found = @book.vba.project.VBComponents.Item(name)
          found if found.Type == USERFORM_TYPE
        rescue WineOLE::RemoteError
          nil
        end
        return nil if component.nil?

        Form.new(name, component, @book)
      end

      # The helper block. `Show 0` is modeless, and it is the only form the
      # wrapper offers: a modal Show blocks Excel's message loop, and with
      # it the bridge -- measured, the bridge freezes until the form closes.
      def self.helper(name)
        <<~VBA
          Function WineOLE_Form_#{name}() As Object
              Set WineOLE_Form_#{name} = #{name}
          End Function
          Sub WineOLE_Show_#{name}()
              #{name}.Show 0
          End Sub
          Sub WineOLE_Hide_#{name}()
              #{name}.Hide
          End Sub
          Sub WineOLE_Unload_#{name}()
              Unload #{name}
          End Sub
        VBA
      end
    end

    # One UserForm. `ole` is the Designer (the design-time form: Caption,
    # Width, Height, and the Controls that `controls` wraps). `instance` is
    # the loaded form -- the object that shows, hides and fires events.
    #
    # Name note (M0): `name`, `component`, `ole`, `instance`, `show`,
    # `hide`, `unload` measured free on the Designer; `controls` is the
    # deliberate shadow.
    class Form
      include Passthrough

      attr_reader :name, :component

      # Writing the helper on every construction (add and re-bind alike) is
      # what makes a form in a reopened workbook showable without the
      # caller remembering to do anything; write is an upsert.
      def initialize(name, component, book)
        @name = name
        @component = component
        @book = book
        @ole = component.Designer
        @runtime = {}
        @instance = nil
        @book.vba.write(Forms.helper(name), name: "form_#{name}")
      end

      def controls
        @controls ||= UserFormControls.new(self, @book.vba)
      end

      # The default instance. Referencing it loads the form if it is not
      # loaded (VBA auto-instantiation), so `shown?` before `show` answers
      # false and leaves the form loaded but hidden. Cached until `unload`.
      def instance
        @instance ||= run('Form')
      end

      # The live counterpart of a design-time control, by name. Cached so
      # that `on` and the `off` that undoes it meet the same Events.
      def runtime_control(control_name)
        @runtime[control_name] ||= instance.Controls.Item(control_name)
      end

      # Modeless, always. Returns as soon as the form is on screen; the
      # bridge stays responsive, which is what lets events reach Ruby.
      def show
        run('Show')
        self
      end

      # Through VBA, like show and unload: the extender's Hide answers
      # GetIDsOfNames but refuses every out-of-process Invoke (measured,
      # DISP_E_MEMBERNOTFOUND whatever the flags), while Visible reads fine.
      def hide
        run('Hide')
        self
      end

      def shown?
        instance.Visible ? true : false
      end

      # Unloading destroys the runtime controls, and with them every event
      # connection Ruby holds on them, so those are closed first -- on our
      # side, deliberately, rather than left to fail when the object is
      # gone. The next `instance` loads a fresh form.
      def unload
        @runtime.each_value { |control| control.ole_events.close }
        @instance&.ole_events&.close
        @runtime.clear
        @instance = nil
        run('Unload')
        self
      end

      private

      # Qualified with the workbook name so the right book's procedure runs
      # when several are open (M2 measured this form working on Excel 11).
      def run(verb)
        @book.ole.Application.Run("'#{@book.ole.Name}'!WineOLE_#{verb}_#{@name}")
      end
    end
  end
end
