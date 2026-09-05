require 'minitest/autorun'
require_relative '../../../lib/wineole/msoffice/controls'
require_relative '../../../lib/wineole/msoffice/sheet'

# A COM Range with a position -- what Sheet#[] hands geometry() via `.ole`.
class FakeComRangeWithBox
  def Left   = 10.5
  def Top    = 20.0
  def Width  = 30.0
  def Height = 40.0
end

# Just enough Worksheet for Sheet#[] to resolve an address: records what
# was asked for.
class FakeComSheetForGeometry
  attr_reader :range_calls

  def initialize = @range_calls = []
  def Name = 'Sheet1'

  def Range(addr)
    @range_calls << addr
    FakeComRangeWithBox.new
  end
end

# A COM collection whose Item(name) knows exactly one name.
class FakeItemHost
  def initialize(known) = @known = known

  def Item(name)
    raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'COM error (0x800A03EC)') unless name == @known

    :found
  end
end

class ControlsModuleTest < Minitest::Test
  C = WineOLE::MSOffice::Controls

  def geometry_sheet
    WineOLE::MSOffice::Sheet.new(FakeComSheetForGeometry.new, version: 11.0)
  end

  # --- kind tables -----------------------------------------------------

  def test_form_kinds_map_to_the_nine_legacy_collections
    assert_equal %w[Buttons CheckBoxes OptionButtons ListBoxes DropDowns Spinners ScrollBars Labels GroupBoxes],
      C::FORM_KINDS.values
    assert_equal 'CheckBoxes', C.form_collection_for(:check_box)
  end

  def test_an_unknown_form_kind_lists_the_valid_ones
    err = assert_raises(ArgumentError) { C.form_collection_for(:command_button) }
    assert_match(/unknown form control kind :command_button/, err.message)
    assert_match(/:button, :check_box, :option_button/, err.message)
    assert_match(/sheet\.activex/, err.message, 'points at the family that has this kind')
  end

  def test_a_string_is_not_a_form_control_kind
    err = assert_raises(ArgumentError) { C.form_collection_for('Forms.CommandButton.1') }
    assert_match(/form controls have no ProgID/, err.message)
  end

  def test_msforms_kinds_map_to_progids
    assert_equal 'Forms.CommandButton.1', C.progid_for(:command_button)
    assert_equal 'Forms.Image.1', C.progid_for(:image)
    assert_equal 11, C::MSFORMS_KINDS.length
  end

  def test_a_string_kind_is_the_progid_verbatim
    assert_equal 'MSComctlLib.ProgCtrl.2', C.progid_for('MSComctlLib.ProgCtrl.2')
  end

  def test_an_unknown_msforms_kind_lists_the_valid_ones
    err = assert_raises(ArgumentError) { C.progid_for(:button) }
    assert_match(/unknown ActiveX kind :button/, err.message)
    assert_match(/:command_button, :text_box/, err.message)
    assert_match(/or pass a ProgID String/, err.message)
  end

  def test_kind_is_recovered_from_a_progid
    assert_equal :toggle_button, C.kind_for_progid('Forms.ToggleButton.1')
    assert_equal 'Other.Ctrl.1', C.kind_for_progid('Other.Ctrl.1')
  end

  def test_form_control_types_skip_the_edit_box
    assert_equal :button, C::FORM_CONTROL_TYPES[0]
    assert_equal :spinner, C::FORM_CONTROL_TYPES[9]
    assert_nil C::FORM_CONTROL_TYPES[3]
  end

  # --- names -----------------------------------------------------------

  def test_a_vba_identifier_is_a_valid_name
    assert_nil C.check_name!('OK_Button2')
    assert_nil C.check_name!('a' * 31)
  end

  def test_names_that_cannot_appear_in_name_click_are_refused
    ['', '1st', 'my button', 'ok-button', 'a' * 32, nil, :sym].each do |bad|
      err = assert_raises(ArgumentError, bad.inspect) { C.check_name!(bad) }
      assert_match(/name:/, err.message)
    end
  end

  def test_event_names_are_vba_identifiers
    assert_nil C.check_event!('KeyDown')
    err = assert_raises(ArgumentError) { C.check_event!('Key Down') }
    assert_match(/event/, err.message)
  end

  def test_check_free_passes_when_the_lookup_raises
    assert_nil C.check_free!(FakeItemHost.new('Other'), 'Fresh', 'sheet')
  end

  def test_check_free_refuses_a_name_the_host_already_has
    err = assert_raises(ArgumentError) { C.check_free!(FakeItemHost.new('Taken'), 'Taken', 'sheet') }
    assert_match(/this sheet already has a control named "Taken"/, err.message)
  end

  # --- properties ------------------------------------------------------

  def test_snake_case_keys_become_pascal_case
    assert_equal 'Caption', C.pascal(:caption)
    assert_equal 'BackColor', C.pascal(:back_color)
    assert_equal 'ListFillRange', C.pascal(:list_fill_range)
    assert_equal 'OnAction', C.pascal(:on_action)
  end

  def test_put_sends_one_property_assignment
    target = Class.new do
      attr_reader :puts_seen
      def initialize = @puts_seen = []
      def method_missing(name, *args) = @puts_seen << [name, args]
      def respond_to_missing?(*) = true
    end.new
    C.put(target, :back_color, 255)
    assert_equal [[:BackColor=, [255]]], target.puts_seen
  end

  def test_the_five_host_props
    assert_equal %i[linked_cell list_fill_range visible print_object placement], C::HOST_PROPS
  end

  # --- geometry --------------------------------------------------------

  def test_at_resolves_through_the_sheet_to_the_ranges_four_values
    sheet = geometry_sheet
    box = C.geometry(sheet: sheet, at: 'B2:C4', left: nil, top: nil, width: nil, height: nil)
    assert_equal [10.5, 20.0, 30.0, 40.0], box
    assert_equal ['B2:C4'], sheet.ole.range_calls
  end

  def test_at_inherits_the_sheets_own_sheet_rule
    err = assert_raises(ArgumentError) do
      C.geometry(sheet: geometry_sheet, at: 'Other!B2', left: nil, top: nil, width: nil, height: nil)
    end
    assert_match(/Other/, err.message)
  end

  def test_four_points_pass_through_unchanged
    box = C.geometry(sheet: geometry_sheet, at: nil, left: 1, top: 2, width: 3, height: 4)
    assert_equal [1, 2, 3, 4], box
  end

  def test_mixing_at_and_points_raises
    err = assert_raises(ArgumentError) do
      C.geometry(sheet: geometry_sheet, at: 'B2', left: 1, top: nil, width: nil, height: nil)
    end
    assert_match(/not both/, err.message)
  end

  def test_no_position_at_all_raises
    err = assert_raises(ArgumentError) do
      C.geometry(sheet: geometry_sheet, at: nil, left: nil, top: nil, width: nil, height: nil)
    end
    assert_match(/no position/, err.message)
  end

  def test_a_partial_box_names_what_is_missing
    err = assert_raises(ArgumentError) do
      C.geometry(sheet: geometry_sheet, at: nil, left: 1, top: 2, width: nil, height: nil)
    end
    assert_match(/missing \[:width, :height\]/, err.message)
  end

  def test_a_userform_has_no_cells
    err = assert_raises(ArgumentError) do
      C.geometry(sheet: nil, at: 'B2', left: nil, top: nil, width: nil, height: nil)
    end
    assert_match(/a UserForm has no cells/, err.message)
  end

  def test_geometry_never_touches_excel_when_it_refuses
    sheet = geometry_sheet
    assert_raises(ArgumentError) { C.geometry(sheet: sheet, at: 'B2', left: 1, top: 2, width: 3, height: 4) }
    assert_empty sheet.ole.range_calls
  end
end

# Records subscriptions the way Events does, without a bridge.
class FakeEventsForControl
  attr_reader :on_calls, :off_calls, :closed

  def initialize
    @on_calls = []
    @off_calls = []
    @closed = false
  end

  def on(name, args: true, &block)
    @on_calls << [name, args, block]
    :subscription
  end

  def off(x) = @off_calls << x
  def close = @closed = true
end

# A COM object that records property puts and hands out one Events.
class FakeComObjectForControl
  attr_reader :puts_seen, :ole_events

  def initialize
    @puts_seen = []
    @ole_events = FakeEventsForControl.new
  end

  def method_missing(name, *args)
    if name.to_s.end_with?('=')
      @puts_seen << [name, args.first]
    else
      super
    end
  end

  def respond_to_missing?(name, _include_private = false) = name.to_s.end_with?('=')

  # What COM's `Object` member answers on an MSForms control: a different
  # object from the control itself (M0).
  def Object = :the_com_member
end

# Records what a VBA writer was asked to write, and where.
class FakeVBAWriter
  attr_reader :writes

  def initialize = @writes = []

  def write(code, name:, into: nil)
    @writes << { code: code, name: name, into: into }
    self
  end
end

# The Form a :userform Control talks to: hands out one runtime control per name.
class FakeFormForControl
  attr_reader :name, :runtime_requests

  def initialize(name)
    @name = name
    @runtime_requests = []
    @runtime = Hash.new { |h, k| h[k] = FakeComObjectForControl.new }
  end

  def runtime_control(control_name)
    @runtime_requests << control_name
    @runtime[control_name]
  end
end

class ControlTest < Minitest::Test
  Control = WineOLE::MSOffice::Control

  def form_control(writer: FakeVBAWriter.new)
    ole = FakeComObjectForControl.new
    Control.new(name: 'Go', kind: :button, family: :form_control, ole: ole, ole_object: ole, vba: writer)
  end

  def activex(writer: FakeVBAWriter.new)
    Control.new(name: 'Go', kind: :command_button, family: :activex,
                ole: FakeComObjectForControl.new, ole_object: FakeComObjectForControl.new, vba: writer)
  end

  def userform(writer: FakeVBAWriter.new, form: FakeFormForControl.new('AppForm'))
    ole = FakeComObjectForControl.new
    Control.new(name: 'OK', kind: :command_button, family: :userform, ole: ole, ole_object: ole,
                vba: writer, form: form)
  end

  # --- readers and passthrough -----------------------------------------

  def test_readers
    ctl = activex
    assert_equal 'Go', ctl.name
    assert_equal :command_button, ctl.kind
    assert_equal :activex, ctl.family
    refute_same ctl.ole, ctl.ole_object, 'a worksheet ActiveX control is two objects'
  end

  def test_unknown_methods_go_to_ole_object_not_ole
    ctl = activex
    ctl.Caption = 'Go'
    assert_equal [[:Caption=, 'Go']], ctl.ole_object.puts_seen
    assert_empty ctl.ole.puts_seen
  end

  def test_ole_object_is_ole_for_the_other_two_families
    plain = form_control
    assert_same plain.ole, plain.ole_object
    ctl = userform
    assert_same ctl.ole, ctl.ole_object
  end

  def test_the_com_member_Object_is_still_reachable
    # The reader is `ole_object`, not `object`, precisely so that a bare
    # `ctl.Object` passes through to the COM member of that name.
    assert_equal :the_com_member, activex.Object
    assert_equal :the_com_member, userform.Object
  end

  # --- Ruby-block handlers ---------------------------------------------

  def test_a_form_control_refuses_a_block
    err = assert_raises(ArgumentError) { form_control.on('Click') { nil } }
    assert_match(/form controls have no COM events/, err.message)
    assert_match(/sheet\.activex/, err.message)
    assert_nil form_control.events
  end

  def test_off_on_a_form_control_raises_the_same_way
    assert_raises(ArgumentError) { form_control.off('Click') }
  end

  def test_activex_on_reaches_the_inner_objects_events
    ctl = activex
    blk = proc { nil }
    assert_equal :subscription, ctl.on('Click', args: false, &blk)
    assert_equal [['Click', false, blk]], ctl.ole_object.ole_events.on_calls
    assert_same ctl.ole_object.ole_events, ctl.events
  end

  def test_activex_off_forwards
    ctl = activex
    ctl.off(:subscription)
    assert_equal [:subscription], ctl.ole_object.ole_events.off_calls
  end

  def test_userform_on_reaches_the_runtime_control_not_the_designer
    form = FakeFormForControl.new('AppForm')
    ctl = userform(form: form)
    ctl.on('Click') { nil }
    assert_equal ['OK'], form.runtime_requests
    assert_equal 1, form.runtime_control('OK').ole_events.on_calls.length
    assert_empty ctl.ole.ole_events.on_calls, 'the design-time control never fires events'
  end

  def test_runtime_is_userform_only
    form = FakeFormForControl.new('AppForm')
    assert_same form.runtime_control('OK'), userform(form: form).runtime
    err = assert_raises(ArgumentError) { activex.runtime }
    assert_match(/only a UserForm control has a runtime instance/, err.message)
  end

  # --- VBA handlers ----------------------------------------------------

  def test_a_form_control_gets_a_click_sub_and_on_action
    writer = FakeVBAWriter.new
    ctl = form_control(writer: writer)
    assert_same ctl, ctl.vba("Range(\"A1\").Value = 1\nBeep")
    assert_equal [{ code: "Sub Go_Click()\n    Range(\"A1\").Value = 1\n    Beep\nEnd Sub", name: 'Go_Click', into: nil }],
      writer.writes
    assert_equal [[:OnAction=, 'Go_Click']], ctl.ole.puts_seen
  end

  def test_a_form_control_takes_no_event_name
    err = assert_raises(ArgumentError) { form_control.vba('Click', 'Beep') }
    assert_match(/fires only Click/, err.message)
  end

  def test_an_activex_handler_is_a_private_sub_in_the_sheet_writer
    writer = FakeVBAWriter.new
    activex(writer: writer).vba('Click', 'Beep')
    assert_equal [{ code: "Private Sub Go_Click()\n    Beep\nEnd Sub", name: 'Go_Click', into: nil }],
      writer.writes
  end

  def test_params_go_into_the_signature_verbatim
    writer = FakeVBAWriter.new
    activex(writer: writer).vba('KeyDown', 'Beep', params: 'ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer')
    assert_equal "Private Sub Go_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)\n    Beep\nEnd Sub",
      writer.writes.first[:code]
    assert_equal 'Go_KeyDown', writer.writes.first[:name]
  end

  def test_a_userform_handler_goes_into_the_form_module
    writer = FakeVBAWriter.new
    userform(writer: writer).vba('Click', 'Beep')
    assert_equal [{ code: "Private Sub OK_Click()\n    Beep\nEnd Sub", name: 'OK_Click', into: 'AppForm' }],
      writer.writes
  end

  def test_activex_and_userform_need_an_event_name
    err = assert_raises(ArgumentError) { activex.vba('Beep') }
    assert_match(/vba\(event, body\)/, err.message)
    assert_raises(ArgumentError) { userform.vba('Beep') }
  end

  def test_an_event_name_that_is_not_an_identifier_is_refused_before_writing
    writer = FakeVBAWriter.new
    assert_raises(ArgumentError) { activex(writer: writer).vba('Key Down', 'Beep') }
    assert_empty writer.writes
  end

  def test_blank_lines_in_the_body_stay_blank
    writer = FakeVBAWriter.new
    activex(writer: writer).vba('Click', "A\n\nB\n")
    assert_equal "Private Sub Go_Click()\n    A\n\n    B\nEnd Sub", writer.writes.first[:code]
  end
end

# --- fakes for the two sheet-side collections ----------------------------
#
# One name registry (a Hash name => shape) backs Shapes, the nine legacy
# collections and OLEObjects, the way one sheet does in Excel: a control
# added through any collection is visible to Shapes.Item and Shapes.Count.
# That is what lets the rollback tests tell "added, then deleted" from
# "never added".

def remote_error(hresult) = WineOLE::RemoteError.new('WIN32OLERuntimeError', "COM error (#{hresult})")

class FakeCodeModuleForControls
  attr_reader :lines

  def initialize = @lines = []
  def CountOfLines = @lines.length
  def Lines(start, count) = @lines[(start - 1), count].join("\n")
  def DeleteLines(start, count) = @lines.slice!((start - 1), count)
  def AddFromString(text) = @lines.concat(text.split(/\r?\n/))
  def text = @lines.join("\n")
end

class FakeVBComponentForControls
  attr_reader :Type, :CodeModule
  attr_accessor :Name

  def initialize(name, type)
    @Name = name
    @Type = type
    @CodeModule = FakeCodeModuleForControls.new
  end
end

class FakeVBComponentsForControls
  def initialize = @components = [FakeVBComponentForControls.new('Sheet1', 100)]

  def Item(name)
    @components.find { |c| c.Name == name } or raise remote_error('0x800A0009')
  end

  def Add(type)
    FakeVBComponentForControls.new("Module#{@components.length}", type).tap { |c| @components << c }
  end

  def Remove(component) = @components.delete(component)
  def names = @components.map(&:Name)
  def module_text(name) = Item(name).CodeModule.text
end

class FakeVBProjectForControls
  attr_reader :VBComponents

  def initialize = @VBComponents = FakeVBComponentsForControls.new
end

class FakeWorkbookForControls
  attr_reader :project

  def initialize(denied:)
    @denied = denied
    @project = FakeVBProjectForControls.new
  end

  def Name = 'Book1'

  def VBProject
    raise remote_error('0x800A03EC') if @denied

    @project
  end
end

# A legacy form control: what Buttons.Add / CheckBoxes.Add etc. hand back.
class FakeFormControlShape
  attr_reader :collection, :box, :puts_seen, :name, :deleted

  def initialize(registry, collection, box, name)
    @registry = registry
    @collection = collection
    @box = box
    @name = name
    @puts_seen = []
    @deleted = false
    @registry[name] = self
  end

  def Name = @name

  # Excel refuses some renames (an invalid or reserved name) with 0x800A03EC;
  # the fake does so for exactly one value so the rollback path can be tested.
  def Name=(value)
    raise remote_error('0x800A03EC') if value == 'Refused'

    @registry.delete(@name)
    @name = value
    @registry[value] = self
    @puts_seen << [:Name=, value]
  end

  # Endless method syntax cannot define a setter, so these two are written
  # as ordinary one-line-body methods instead.
  def Caption=(value)
    @puts_seen << [:Caption=, value]
  end

  def OnAction=(value)
    @puts_seen << [:OnAction=, value]
  end

  def FormControlType
    kind = WineOLE::MSOffice::Controls::FORM_KINDS.key(@collection)
    WineOLE::MSOffice::Controls::FORM_CONTROL_TYPES.key(kind)
  end

  def Delete
    @registry.delete(@name)
    @deleted = true
  end

  # Any other property put is a member Excel does not have.
  def method_missing(name, *args)
    raise remote_error('0x80020006') if name.to_s.end_with?('=')

    super
  end

  def respond_to_missing?(name, _include_private = false) = name.to_s.end_with?('=')
end

class FakeLegacyCollection
  def initialize(registry, collection, add_calls)
    @registry = registry
    @collection = collection
    @add_calls = add_calls
  end

  def Add(left, top, width, height)
    @add_calls << [@collection, left, top, width, height]
    FakeFormControlShape.new(@registry, @collection, [left, top, width, height], "#{@collection} #{@registry.length + 1}")
  end

  def Item(name)
    shape = @registry[name]
    raise remote_error('0x800A03EC') unless shape.is_a?(FakeFormControlShape) && shape.collection == @collection

    shape
  end
end

# The MSForms control inside an OLEObject.
class FakeMSFormsObjectForControls
  attr_reader :puts_seen, :ole_events

  def initialize(progid)
    @progid = progid
    @puts_seen = []
    @ole_events = FakeEventsForControl.new
  end

  def method_missing(name, *args)
    return super unless name.to_s.end_with?('=')
    # An Image has no Caption; Excel refuses the put.
    raise remote_error('0x80020006') if @progid == 'Forms.Image.1' && name == :Caption=

    @puts_seen << [name, args.first]
  end

  def respond_to_missing?(name, _include_private = false) = name.to_s.end_with?('=')
end

class FakeOLEObjectForControls
  attr_reader :progID, :box, :puts_seen, :name, :deleted, :inner

  HOST_PUTS = %i[Name= LinkedCell= ListFillRange= Visible= PrintObject= Placement=].freeze

  def initialize(registry, progid, box, name)
    @registry = registry
    @progID = progid
    @box = box
    @name = name
    @puts_seen = []
    @deleted = false
    @inner = FakeMSFormsObjectForControls.new(progid)
    @registry[name] = self
  end

  def Name = @name
  def Object = @inner

  # Same refused rename as FakeFormControlShape, for the same reason.
  def Name=(value)
    raise remote_error('0x800A03EC') if value == 'Refused'

    @registry.delete(@name)
    @name = value
    @registry[value] = self
    @puts_seen << [:Name=, value]
  end

  def Delete
    @registry.delete(@name)
    @deleted = true
  end

  # Shape.FormControlType on an ActiveX shape is a COM error in Excel.
  def FormControlType = raise(remote_error('0x800A03EC'))

  def method_missing(name, *args)
    return super unless name.to_s.end_with?('=')
    raise remote_error('0x80020006') unless HOST_PUTS.include?(name)

    @puts_seen << [name, args.first]
  end

  def respond_to_missing?(name, _include_private = false) = name.to_s.end_with?('=')
end

class FakeOLEObjectsForControls
  def initialize(registry, add_calls)
    @registry = registry
    @add_calls = add_calls
  end

  # The passthrough trap, reproduced: Excel 11 fails OLEObjects.Add with
  # 0x800A03EC unless all four of Left/Top/Width/Height come with ClassType.
  def Add(**named)
    @add_calls << named
    unless %i[ClassType Left Top Width Height].all? { |k| named.key?(k) }
      raise remote_error('0x800A03EC')
    end

    box = named.values_at(:Left, :Top, :Width, :Height)
    FakeOLEObjectForControls.new(@registry, named[:ClassType], box, "OLE#{@registry.length + 1}")
  end

  def Item(name)
    shape = @registry[name]
    raise remote_error('0x800A03EC') unless shape.is_a?(FakeOLEObjectForControls)

    shape
  end
end

class FakeShapesForControls
  def initialize(registry) = @registry = registry
  def Count = @registry.length

  def Item(name)
    @registry[name] or raise remote_error('0x800A03EC')
  end
end

class FakeComSheetForControls
  attr_reader :registry, :add_calls, :ole_add_calls, :range_calls, :parent

  def initialize(denied: false)
    @registry = {}
    @add_calls = []
    @ole_add_calls = []
    @range_calls = []
    @parent = FakeWorkbookForControls.new(denied: denied)
  end

  def Name = 'Sheet1'
  def CodeName = 'Sheet1'
  def Parent = @parent
  def Shapes = FakeShapesForControls.new(@registry)
  def OLEObjects = FakeOLEObjectsForControls.new(@registry, @ole_add_calls)

  def Range(addr)
    @range_calls << addr
    FakeComRangeWithBox.new
  end

  WineOLE::MSOffice::Controls::FORM_KINDS.each_value do |collection|
    define_method(collection) { FakeLegacyCollection.new(@registry, collection, @add_calls) }
  end
end

class FormControlsTest < Minitest::Test
  BOX = { left: 1, top: 2, width: 3, height: 4 }.freeze

  def sheet(denied: false)
    WineOLE::MSOffice::Sheet.new(FakeComSheetForControls.new(denied: denied), version: 11.0)
  end

  def test_readers_are_memoized_and_distinct
    s = sheet
    assert_same s.form_controls, s.form_controls
    assert_same s.activex, s.activex
    refute_same s.form_controls, s.activex
  end

  def test_add_places_through_the_kinds_collection_at_the_given_box
    s = sheet
    ctl = s.form_controls.add(:button, name: 'Go', **BOX)
    assert_equal :form_control, ctl.family
    assert_equal :button, ctl.kind
    assert_equal 'Go', ctl.name
    assert_equal [['Buttons', 1, 2, 3, 4]], s.ole.add_calls
    assert_equal ['Go'], s.ole.registry.keys
    assert_same ctl.ole, ctl.ole_object
  end

  def test_add_with_at_reads_the_ranges_box
    s = sheet
    ctl = s.form_controls.add(:check_box, name: 'Agree', at: 'B2')
    assert_equal ['B2'], s.ole.range_calls
    assert_equal [10.5, 20.0, 30.0, 40.0], ctl.ole.box
    assert_equal 'CheckBoxes', ctl.ole.collection
  end

  def test_rename_comes_before_the_properties
    ctl = sheet.form_controls.add(:button, name: 'Go', caption: 'Run', **BOX)
    assert_equal [[:Name=, 'Go'], [:Caption=, 'Run']], ctl.ole.puts_seen
  end

  def test_a_refused_property_deletes_the_control_and_reraises
    s = sheet
    err = assert_raises(WineOLE::RemoteError) do
      s.form_controls.add(:button, name: 'Go', bogus: 1, **BOX)
    end
    assert_match(/0x80020006/, err.message)
    assert_equal 1, s.ole.add_calls.length, 'the control was added'
    assert_equal 0, s.ole.Shapes.Count, 'and then deleted'
  end

  def test_a_refused_rename_deletes_the_control_and_reraises
    s = sheet
    err = assert_raises(WineOLE::RemoteError) do
      s.form_controls.add(:button, name: 'Refused', caption: 'never set', **BOX)
    end
    assert_match(/0x800A03EC/, err.message)
    assert_equal 1, s.ole.add_calls.length, 'the control was added'
    assert_equal 0, s.ole.Shapes.Count, 'and then deleted'
  end

  def test_a_duplicate_name_is_refused_before_excel_is_asked_again
    s = sheet
    s.form_controls.add(:button, name: 'Go', **BOX)
    err = assert_raises(ArgumentError) { s.form_controls.add(:label, name: 'Go', **BOX) }
    assert_match(/already has a control named "Go"/, err.message)
    assert_equal 1, s.ole.add_calls.length
  end

  def test_argument_checks_run_before_excel_is_touched
    s = sheet
    assert_raises(ArgumentError) { s.form_controls.add(:command_button, name: 'Go', **BOX) }
    assert_raises(ArgumentError) { s.form_controls.add(:button, name: 'bad name', **BOX) }
    assert_raises(ArgumentError) { s.form_controls.add(:button, name: 'Go', left: 1) }
    assert_raises(ArgumentError) { s.form_controls.add(:button, name: 'Go', at: 'B2', left: 1) }
    assert_empty s.ole.add_calls
    assert_empty s.ole.range_calls
  end

  def test_lookup_rebinds_by_form_control_type
    s = sheet
    added = s.form_controls.add(:spinner, name: 'Spin', **BOX)
    found = s.form_controls['Spin']
    assert_equal :spinner, found.kind
    assert_equal :form_control, found.family
    assert_same added.ole, found.ole
  end

  def test_lookup_is_nil_for_a_missing_name_and_for_an_activex_shape
    s = sheet
    assert_nil s.form_controls['Nope']
    s.activex.add(:command_button, name: 'X', **BOX)
    assert_nil s.form_controls['X']
  end

  def test_a_vba_handler_lands_in_the_wrapper_module_of_the_parent_workbook
    s = sheet
    ctl = s.form_controls.add(:button, name: 'Go', **BOX)
    ctl.vba('Beep')
    components = s.ole.parent.project.VBComponents
    assert_includes components.names, 'WineOLE'
    assert_match(/^Sub Go_Click\(\)$/, components.module_text('WineOLE'))
    assert_equal '', components.module_text('Sheet1'), 'not the sheet module'
    assert_includes ctl.ole.puts_seen, [:OnAction=, 'Go_Click']
  end

  def test_a_denied_project_is_access_denied_not_a_remote_error
    ctl = sheet(denied: true).form_controls.add(:button, name: 'Go', **BOX)
    assert_raises(WineOLE::MSOffice::VBA::AccessDenied) { ctl.vba('Beep') }
  end
end

class ActiveXControlsTest < Minitest::Test
  BOX = { left: 1, top: 2, width: 3, height: 4 }.freeze

  def sheet = WineOLE::MSOffice::Sheet.new(FakeComSheetForControls.new, version: 11.0)

  def test_add_passes_classtype_and_all_four_points_by_name
    s = sheet
    ctl = s.activex.add(:command_button, name: 'Go', **BOX)
    assert_equal [{ ClassType: 'Forms.CommandButton.1', Left: 1, Top: 2, Width: 3, Height: 4 }], s.ole.ole_add_calls
    assert_equal :activex, ctl.family
    assert_equal :command_button, ctl.kind
    assert_same ctl.ole.inner, ctl.ole_object
    assert_equal ['Go'], s.ole.registry.keys
  end

  def test_a_string_kind_is_the_progid_and_stays_the_kind
    s = sheet
    ctl = s.activex.add('Other.Ctrl.1', name: 'Ext', **BOX)
    assert_equal 'Other.Ctrl.1', s.ole.ole_add_calls.first[:ClassType]
    assert_equal 'Other.Ctrl.1', ctl.kind
  end

  def test_properties_are_routed_to_the_host_or_the_inner_object
    ctl = sheet.activex.add(:command_button, name: 'Go', caption: 'Go', linked_cell: 'A1', back_color: 255, **BOX)
    assert_equal [[:Name=, 'Go'], [:LinkedCell=, 'A1']], ctl.ole.puts_seen
    assert_equal [[:Caption=, 'Go'], [:BackColor=, 255]], ctl.ole_object.puts_seen
  end

  def test_a_refused_inner_property_deletes_the_control_and_reraises
    s = sheet
    err = assert_raises(WineOLE::RemoteError) { s.activex.add(:image, name: 'Pic', caption: 'x', **BOX) }
    assert_match(/0x80020006/, err.message)
    assert_equal 1, s.ole.ole_add_calls.length
    assert_equal 0, s.ole.Shapes.Count
  end

  def test_a_refused_rename_deletes_the_control_and_reraises
    s = sheet
    err = assert_raises(WineOLE::RemoteError) { s.activex.add(:command_button, name: 'Refused', **BOX) }
    assert_match(/0x800A03EC/, err.message)
    assert_equal 1, s.ole.ole_add_calls.length
    assert_equal 0, s.ole.Shapes.Count
  end

  def test_a_name_taken_by_a_form_control_is_refused
    s = sheet
    s.form_controls.add(:button, name: 'Twice', **BOX)
    err = assert_raises(ArgumentError) { s.activex.add(:command_button, name: 'Twice', **BOX) }
    assert_match(/already has a control named "Twice"/, err.message)
    assert_empty s.ole.ole_add_calls
  end

  def test_argument_checks_run_before_excel_is_touched
    s = sheet
    assert_raises(ArgumentError) { s.activex.add(:button, name: 'Go', **BOX) }
    assert_raises(ArgumentError) { s.activex.add(:command_button, name: '1Go', **BOX) }
    assert_raises(ArgumentError) { s.activex.add(:command_button, name: 'Go', width: 3) }
    assert_empty s.ole.ole_add_calls
  end

  def test_lookup_rebinds_and_recovers_the_kind_from_the_progid
    s = sheet
    s.activex.add(:toggle_button, name: 'T', **BOX)
    found = s.activex['T']
    assert_equal :toggle_button, found.kind
    assert_equal :activex, found.family
    assert_nil s.activex['nope']
  end

  def test_lookup_is_nil_for_a_form_control_shape
    s = sheet
    s.form_controls.add(:button, name: 'Plain', **BOX)
    assert_nil s.activex['Plain']
  end

  def test_a_vba_handler_lands_in_the_sheets_own_module
    s = sheet
    s.activex.add(:command_button, name: 'Go', **BOX).vba('Click', 'Beep')
    components = s.ole.parent.project.VBComponents
    assert_match(/^Private Sub Go_Click\(\)$/, components.module_text('Sheet1'))
    refute_includes components.names, 'WineOLE'
  end
end
