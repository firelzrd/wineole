require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/forms'
require_relative '../../../lib/wineole/msoffice/book'

def remote_error_for_forms(hresult) = WineOLE::RemoteError.new('WIN32OLERuntimeError', "COM error (#{hresult})")

class FakeEventsForForm
  attr_reader :on_calls, :closed

  def initialize
    @on_calls = []
    @closed = 0
  end

  def on(name, args: true, &block)
    @on_calls << [name, args, block]
    :subscription
  end

  def off(_x) = nil
  def close = @closed += 1
end

# A control on the loaded form: fires events, has live values.
class FakeRuntimeControl
  attr_reader :ole_events, :puts_seen

  def initialize
    @ole_events = FakeEventsForForm.new
    @puts_seen = []
  end

  def method_missing(name, *args)
    return super unless name.to_s.end_with?('=')

    @puts_seen << [name, args.first]
  end

  def respond_to_missing?(name, _include_private = false) = name.to_s.end_with?('=')
end

class FakeRuntimeControls
  def initialize = @controls = Hash.new { |h, k| h[k] = FakeRuntimeControl.new }
  def Item(name) = @controls[name]
end

# What WineOLE_Form_<Name> hands back: the form's default instance.
class FakeUserFormInstance
  attr_reader :ole_events, :Controls, :calls

  def initialize
    @ole_events = FakeEventsForForm.new
    @Controls = FakeRuntimeControls.new
    @calls = []
    @visible = false
  end

  def shown! = @visible = true
  def Visible = @visible
end

# A design-time MSForms control on the Designer.
class FakeDesignControl
  attr_reader :Name, :puts_seen

  def initialize(name)
    @Name = name
    @puts_seen = []
  end

  def method_missing(name, *args)
    return super unless name.to_s.end_with?('=')
    raise remote_error_for_forms('0x80020006') if name == :Bogus=

    @puts_seen << [name, args.first]
  end

  def respond_to_missing?(name, _include_private = false) = name.to_s.end_with?('=')
end

class FakeDesignerControls
  attr_reader :add_calls

  def initialize
    @controls = {}
    @add_calls = []
  end

  def Add(progid, name)
    @add_calls << [progid, name]
    @controls[name] = FakeDesignControl.new(name)
  end

  def Item(name)
    @controls[name] or raise remote_error_for_forms('0x800A03EC')
  end

  def Remove(name) = @controls.delete(name)
  def Count = @controls.length
end

class FakeDesigner
  attr_reader :Controls

  def initialize = @Controls = FakeDesignerControls.new
end

class FakeVBComponentForForms
  attr_reader :Name, :Type, :Designer

  def initialize(name, type)
    @Name = name
    @Type = type
    @Designer = FakeDesigner.new
  end
end

class FakeVBComponentsForForms
  def initialize(components) = @components = components

  def Item(name)
    @components[name] or raise remote_error_for_forms('0x800A0009')
  end
end

class FakeVBProjectForForms
  attr_reader :VBComponents

  def initialize(components) = @VBComponents = FakeVBComponentsForForms.new(components)
end

# The BookVBA a Form talks to, as a recorder.
class FakeBookVBAForForms
  attr_reader :writes, :components, :removed
  attr_accessor :fail_write_with

  def initialize
    @writes = []
    @components = { 'Module1' => FakeVBComponentForForms.new('Module1', 1) }
    @removed = []
  end

  # `fail_write_with:` lets a test make the helper write fail the way a
  # refused Excel call would, so Forms#add's rollback can be exercised.
  def write(code, name: 'main', into: nil)
    raise @fail_write_with if @fail_write_with

    @writes << { code: code, name: name, into: into }
    self
  end

  def add_component(name, kind: :standard)
    if @components.key?(name)
      raise ArgumentError, "this workbook already has a VBA component named #{name.inspect}"
    end

    @components[name] = FakeVBComponentForForms.new(name, { standard: 1, class: 2, form: 3 }.fetch(kind))
  end

  def remove_component(name)
    @removed << name
    @components.delete(name)
  end

  def project = FakeVBProjectForForms.new(@components)
end

class FakeApplicationForForms
  attr_reader :runs, :instances

  def initialize
    @runs = []
    @instances = []
  end

  # Every Run of the Form function hands back a NEW instance, so a test can
  # tell "cached" from "fetched again".
  def Run(macro)
    @runs << macro
    return nil unless macro.include?('WineOLE_Form_')

    FakeUserFormInstance.new.tap { |i| @instances << i }
  end
end

class FakeWorkbookForForms
  attr_reader :Application

  def initialize = @Application = FakeApplicationForForms.new
  def Name = 'Book1'
end

class FakeBookForForms
  attr_reader :ole, :vba

  def initialize
    @ole = FakeWorkbookForForms.new
    @vba = FakeBookVBAForForms.new
  end
end

class FormsTest < Minitest::Test
  Forms = WineOLE::MSOffice::Forms
  Form = WineOLE::MSOffice::Form

  def setup
    @book = FakeBookForForms.new
    @forms = Forms.new(@book)
  end

  def test_add_makes_a_form_component_and_wraps_its_designer
    form = @forms.add('AppForm')
    assert_instance_of Form, form
    assert_equal 'AppForm', form.name
    assert_equal 3, form.component.Type
    assert_same form.component.Designer, form.ole
    assert_same @book.vba.components['AppForm'], form.component
  end

  def test_add_checks_the_name_before_touching_vba
    err = assert_raises(ArgumentError) { @forms.add('App Form') }
    assert_match(/name:/, err.message)
    refute @book.vba.components.key?('App Form')
  end

  def test_add_writes_the_helper_block_into_the_wrapper_module
    @forms.add('AppForm')
    assert_equal [{ code: Forms.helper('AppForm'), name: 'form_AppForm', into: nil }], @book.vba.writes
  end

  def test_the_helper_text
    expected = <<~VBA
      Function WineOLE_Form_AppForm() As Object
          Set WineOLE_Form_AppForm = AppForm
      End Function
      Sub WineOLE_Show_AppForm()
          AppForm.Show 0
      End Sub
      Sub WineOLE_Hide_AppForm()
          AppForm.Hide
      End Sub
      Sub WineOLE_Unload_AppForm()
          Unload AppForm
      End Sub
    VBA
    assert_equal expected, Forms.helper('AppForm')
  end

  def test_a_taken_name_is_add_components_refusal
    @forms.add('AppForm')
    err = assert_raises(ArgumentError) { @forms.add('AppForm') }
    assert_match(/already has a VBA component named "AppForm"/, err.message)
  end

  def test_a_refused_helper_write_removes_the_component_and_reraises
    error = remote_error_for_forms('0x800A03EC')
    @book.vba.fail_write_with = error

    raised = assert_raises(WineOLE::RemoteError) { @forms.add('AppForm') }

    assert_same error, raised
    refute @book.vba.components.key?('AppForm')
    assert_equal ['AppForm'], @book.vba.removed
  end

  def test_lookup_rebinds_a_userform_and_nothing_else
    @forms.add('AppForm')
    found = @forms['AppForm']
    assert_instance_of Form, found
    assert_same @book.vba.components['AppForm'], found.component
    assert_nil @forms['Module1'], 'a standard module is not a form'
    assert_nil @forms['Nope']
  end

  def test_lookup_rewrites_the_helper_so_a_reopened_workbook_can_show_the_form
    @forms.add('AppForm')
    @book.vba.writes.clear
    @forms['AppForm']
    assert_equal ['form_AppForm'], @book.vba.writes.map { |w| w[:name] }
  end

  # Only the lookup answers "not there" with nil. A refused helper write is
  # a real failure -- a locked project, a refused AddFromString -- and
  # swallowing it would hand back a Form-shaped nil for a form that exists.
  def test_lookup_raises_when_the_helper_write_is_refused
    @forms.add('AppForm')
    error = remote_error_for_forms('0x800A03EC')
    @book.vba.fail_write_with = error

    raised = assert_raises(WineOLE::RemoteError) { @forms['AppForm'] }

    assert_same error, raised
    assert_nil @forms['Nope'], 'a missing name is still nil'
  end

  def test_book_forms_is_memoized
    client = Class.new { def loopback? = true }.new
    book = WineOLE::MSOffice::Book.new(FakeWorkbookForForms.new, client: client, version: 11.0)
    assert_instance_of Forms, book.forms
    assert_same book.forms, book.forms
  end
end

class FormTest < Minitest::Test
  def setup
    @book = FakeBookForForms.new
    @form = WineOLE::MSOffice::Forms.new(@book).add('AppForm')
    @app = @book.ole.Application
  end

  def test_show_runs_the_show_helper_qualified_with_the_workbook_name
    assert_same @form, @form.show
    assert_equal ["'Book1'!WineOLE_Show_AppForm"], @app.runs
  end

  def test_instance_is_fetched_once_through_the_form_function
    first = @form.instance
    assert_same first, @form.instance
    assert_equal ["'Book1'!WineOLE_Form_AppForm"], @app.runs
    assert_same @app.instances.first, first
  end

  def test_hide_runs_the_hide_helper_and_shown_reads_the_instance
    @form.instance.shown!
    assert @form.shown?
    assert_same @form, @form.hide
    assert_equal ["'Book1'!WineOLE_Form_AppForm", "'Book1'!WineOLE_Hide_AppForm"], @app.runs
    assert_empty @form.instance.calls, 'hide must not call the instance: its Hide is not dispatchable'
  end

  def test_runtime_controls_are_cached_per_name
    a = @form.runtime_control('OK')
    assert_same a, @form.runtime_control('OK')
    refute_same a, @form.runtime_control('Cancel')
    assert_instance_of FakeRuntimeControl, a
  end

  def test_unload_closes_every_events_object_before_running_unload_and_forgets_the_instance
    instance = @form.instance
    ok = @form.runtime_control('OK')
    ok.ole_events.on('Click') { nil }

    assert_same @form, @form.unload

    assert_equal 1, ok.ole_events.closed
    assert_equal 1, instance.ole_events.closed
    assert_equal ["'Book1'!WineOLE_Form_AppForm", "'Book1'!WineOLE_Unload_AppForm"], @app.runs
    refute_same instance, @form.instance, 'the next instance is fetched afresh'
    refute_same ok, @form.runtime_control('OK')
  end

  def test_unload_without_an_instance_just_runs_the_helper
    @form.unload
    assert_equal ["'Book1'!WineOLE_Unload_AppForm"], @app.runs
  end

  def test_unknown_methods_reach_the_designer
    designer = @form.ole
    def designer.Width = 240
    assert_equal 240, @form.Width
  end
end

class UserFormControlsTest < Minitest::Test
  BOX = { left: 1, top: 2, width: 3, height: 4 }.freeze

  def setup
    @book = FakeBookForForms.new
    @form = WineOLE::MSOffice::Forms.new(@book).add('AppForm')
    @designer_controls = @form.ole.Controls
  end

  def test_controls_is_memoized
    assert_same @form.controls, @form.controls
    assert_instance_of WineOLE::MSOffice::UserFormControls, @form.controls
  end

  def test_add_names_the_control_at_add_time_then_sizes_it
    ctl = @form.controls.add(:command_button, name: 'OK', **BOX)
    assert_equal [['Forms.CommandButton.1', 'OK']], @designer_controls.add_calls
    assert_equal [[:Left=, 1], [:Top=, 2], [:Width=, 3], [:Height=, 4]], ctl.ole.puts_seen
    assert_equal :userform, ctl.family
    assert_equal :command_button, ctl.kind
    assert_same ctl.ole, ctl.ole_object
  end

  def test_properties_follow_the_box
    ctl = @form.controls.add(:command_button, name: 'OK', caption: 'OK', **BOX)
    assert_equal [:Caption=, 'OK'], ctl.ole.puts_seen.last
  end

  def test_at_is_refused_on_a_form
    err = assert_raises(ArgumentError) { @form.controls.add(:command_button, name: 'OK', at: 'B2') }
    assert_match(/a UserForm has no cells/, err.message)
    assert_empty @designer_controls.add_calls
  end

  def test_a_refused_property_removes_the_control_and_reraises
    assert_raises(WineOLE::RemoteError) { @form.controls.add(:command_button, name: 'OK', bogus: 1, **BOX) }
    assert_equal 1, @designer_controls.add_calls.length
    assert_equal 0, @designer_controls.Count
  end

  def test_a_duplicate_name_is_refused_before_add
    @form.controls.add(:command_button, name: 'OK', **BOX)
    err = assert_raises(ArgumentError) { @form.controls.add(:label, name: 'OK', **BOX) }
    assert_match(/this UserForm already has a control named "OK"/, err.message)
    assert_equal 1, @designer_controls.add_calls.length
  end

  def test_lookup_rebinds_with_no_kind
    @form.controls.add(:text_box, name: 'Name', **BOX)
    found = @form.controls['Name']
    assert_equal :userform, found.family
    assert_nil found.kind
    assert_same @designer_controls.Item('Name'), found.ole
    assert_nil @form.controls['Nope']
  end

  def test_a_block_handler_lands_on_the_runtime_control
    ctl = @form.controls.add(:command_button, name: 'OK', **BOX)
    ctl.on('Click') { nil }
    assert_equal 1, @form.runtime_control('OK').ole_events.on_calls.length
  end

  def test_a_vba_handler_goes_into_the_form_module
    @form.controls.add(:command_button, name: 'OK', **BOX).vba('Click', 'Beep')
    handler = @book.vba.writes.last
    assert_equal({ code: "Private Sub OK_Click()\n    Beep\nEnd Sub", name: 'OK_Click', into: 'AppForm' }, handler)
  end
end
