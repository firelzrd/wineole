require 'minitest/autorun'
require 'timeout'
require_relative '../../lib/wineole/msoffice'
require_relative '../support/excel_integration_helper'

# The controls wrapper against real Excel under Wine: the three families
# place where asked, the two handler paths run, and the passthrough trap
# the wrapper exists to remove is still a trap on the passthrough. Every
# test gets its own Excel; ExcelIntegrationHelper's teardown (the spec's
# test 10) kills only what the test started.
#
# A click is delivered without a UI by setting an MSForms CommandButton's
# Value to true, which fires its Click event.
class WineOLEMSOfficeControlsIntegrationTest < Minitest::Test
  include ExcelIntegrationHelper

  BOX = { left: 10, top: 10, width: 100, height: 30 }.freeze

  # A ProgID outside the wrapper's table that places on a worksheet and on a
  # UserForm on this host (measured 2026-09-04: the Windows Common Controls
  # progress bar). The environment override is for re-measuring on another
  # host; `none` skips the worksheet half.
  M3_SHEET_PROGID = ENV.fetch('WINEOLE_M3_SHEET_PROGID', 'MSComctlLib.ProgCtrl.2').then { |v| v == 'none' ? nil : v }
  M3_FORM_PROGID = ENV.fetch('WINEOLE_M3_FORM_PROGID', 'MSComctlLib.ProgCtrl.2')

  # spec test 1: every kind of every family, at the box it was given, with
  # a caption where the kind has one; timings per family printed for the
  # README.
  def test_every_kind_of_all_three_families_places_where_asked
    with_excel do |xl|
      sheet = xl[':first!']
      form = xl['[]'].forms.add('KindsForm')

      timed('form_controls') do
        WineOLE::MSOffice::Controls::FORM_KINDS.each_key do |kind|
          props = kind == :button ? { caption: 'Go' } : {}
          ctl = sheet.form_controls.add(kind, name: "F#{kind}", **BOX, **props)
          assert_equal "F#{kind}", ctl.ole.Name
          assert_box ctl.ole, BOX
          assert_equal 'Go', ctl.Caption if kind == :button
        end
      end

      timed('activex') do
        WineOLE::MSOffice::Controls::MSFORMS_KINDS.each_key do |kind|
          props = kind == :command_button ? { caption: 'Go' } : {}
          ctl = sheet.activex.add(kind, name: "X#{kind}", **BOX, **props)
          assert_equal "X#{kind}", ctl.ole.Name
          assert_box ctl.ole, BOX
          assert_equal 'Go', ctl.Caption if kind == :command_button
        end
      end

      timed('userform') do
        WineOLE::MSOffice::Controls::MSFORMS_KINDS.each_key do |kind|
          props = kind == :command_button ? { caption: 'Go' } : {}
          ctl = form.controls.add(kind, name: "U#{kind}", **BOX, **props)
          assert_equal "U#{kind}", ctl.ole.Name
          assert_box ctl.ole, BOX
          assert_equal 'Go', ctl.Caption if kind == :command_button
        end
      end

      # Re-binding what was just placed, and a name that is not there.
      assert_equal :button, sheet.form_controls['Fbutton'].kind
      assert_equal :command_button, sheet.activex['Xcommand_button'].kind
      assert_equal 'Ucommand_button', form.controls['Ucommand_button'].ole.Name
      assert_nil sheet.form_controls['nope']
      assert_nil sheet.activex['nope']
      assert_nil form.controls['nope']
    end
  end

  # spec test 2: at: reads the range's box.
  def test_at_matches_the_ranges_box
    with_excel do |xl|
      sheet = xl[':first!']
      range = sheet['B2:C4'].ole
      ctl = sheet.activex.add(:text_box, name: 'AtBox', at: 'B2:C4')
      assert_in_delta range.Left, ctl.ole.Left, 0.5
      assert_in_delta range.Top, ctl.ole.Top, 0.5
      assert_in_delta range.Width, ctl.ole.Width, 0.5
      assert_in_delta range.Height, ctl.ole.Height, 0.5
    end
  end

  # spec test 3: the trap stays measured. OLEObjects.Add with only Left and
  # Top fails on Excel 11; the wrapper, which always sends all four, does
  # not.
  def test_the_passthrough_trap_and_the_wrapper_beside_it
    with_excel do |xl|
      sheet = xl[':first!']
      err = assert_raises(WineOLE::RemoteError) do
        sheet.ole.OLEObjects.Add(ClassType: 'Forms.CommandButton.1', Left: 10, Top: 10)
      end
      assert_match(/0x800A03EC/, err.message)

      ctl = sheet.activex.add(:command_button, name: 'Wrapped', at: 'B2')
      assert_equal 'Wrapped', ctl.ole.Name
    end
  end

  # spec test 4 (M1): a worksheet ActiveX Click reaches a Ruby block.
  def test_a_worksheet_activex_click_reaches_a_ruby_block
    with_excel do |xl|
      sheet = xl[':first!']
      ctl = sheet.activex.add(:command_button, name: 'Go', **BOX)
      got = Queue.new
      ctl.on('Click') { got << :clicked }
      ctl.Value = true
      assert_equal :clicked, wait_for(got, 'the Click callback')
      ctl.off('Click')
    end
  end

  # spec test 5: a VBA handler in the sheet module runs on Click.
  def test_a_worksheet_activex_vba_handler_runs_when_clicked
    with_excel do |xl|
      sheet = xl[':first!']
      ctl = sheet.activex.add(:command_button, name: 'Go', **BOX)
      ctl.vba('Click', 'Range("Z1").Value = 1')
      ctl.Value = true
      assert settled { sheet['Z1'].ole.Value == 1 }, 'Z1 never became 1: the sheet-module handler did not run'
    end
  end

  # spec test 6: a form control's macro is bound through OnAction and runs.
  def test_a_form_control_macro_is_bound_through_on_action
    with_excel do |xl|
      sheet = xl[':first!']
      ctl = sheet.form_controls.add(:button, name: 'Plain', **BOX)
      ctl.vba('Range("Z2").Value = 2')
      assert_match(/Plain_Click\z/, ctl.ole.OnAction)
      xl.ole.Run(ctl.ole.OnAction)
      assert_equal 2, sheet['Z2'].ole.Value
    end
  end

  # spec test 7 (M2): a UserForm button's Click reaches Ruby, the
  # subscription survives hide/show, unload closes it without an error.
  def test_a_userform_button_click_reaches_ruby_across_hide_and_show
    with_excel do |xl|
      form = xl['[]'].forms.add('AppForm')
      ok = form.controls.add(:command_button, name: 'OK', caption: 'OK', **BOX)
      got = Queue.new
      ok.on('Click') { got << :clicked }

      form.show
      assert form.shown?
      ok.runtime.Value = true
      assert_equal :clicked, wait_for(got, 'the first Click')

      form.hide
      refute form.shown?
      form.show
      ok.runtime.Value = true
      assert_equal :clicked, wait_for(got, 'the Click after hide and show')

      form.unload
    end
  end

  # spec test 8 (M3): a ProgID String places a control outside the table,
  # on the sheet when this host has one that goes there, on a UserForm
  # otherwise.
  def test_a_progid_string_places_a_control_outside_the_table
    with_excel do |xl|
      if M3_SHEET_PROGID
        sheet = xl[':first!']
        ctl = sheet.activex.add(M3_SHEET_PROGID, name: 'Ext', **BOX)
        assert_equal M3_SHEET_PROGID, ctl.kind
        assert_equal 'Ext', ctl.ole.Name
        assert_equal M3_SHEET_PROGID, sheet.activex['Ext'].kind
      else
        form = xl['[]'].forms.add('ExtForm')
        ctl = form.controls.add(M3_FORM_PROGID, name: 'Ext', **BOX)
        assert_equal M3_FORM_PROGID, ctl.kind
        assert_equal 'Ext', ctl.ole.Name
        assert_equal 'Ext', form.controls['Ext'].ole.Name
      end
    end
  end

  # spec test 9: a duplicate name is refused across families, and a
  # property Excel rejects rolls the placement back.
  def test_names_are_refused_across_families_and_a_refused_property_leaves_nothing
    with_excel do |xl|
      sheet = xl[':first!']
      sheet.form_controls.add(:button, name: 'Twice', **BOX)
      err = assert_raises(ArgumentError) { sheet.activex.add(:command_button, name: 'Twice', **BOX) }
      assert_match(/already has a control named "Twice"/, err.message)

      before = sheet.ole.Shapes.Count
      assert_raises(WineOLE::RemoteError) { sheet.activex.add(:image, name: 'Pic', caption: 'no', **BOX) }
      assert_equal before, sheet.ole.Shapes.Count, 'the refused control must be gone'
    end
  end

  private

  def with_excel
    with_bridge do |client|
      WineOLE::MSOffice::Excel.run(:create, client: client) do |xl|
        xl.hide
        xl.no_alert do
          xl.ole.Workbooks.Add
          yield xl
        end
      end
    end
  end

  def assert_box(ole, box)
    assert_in_delta box[:left], ole.Left, 0.5
    assert_in_delta box[:top], ole.Top, 0.5
    assert_in_delta box[:width], ole.Width, 0.5
    assert_in_delta box[:height], ole.Height, 0.5
  end

  def wait_for(queue, what)
    Timeout.timeout(30) { queue.pop }
  rescue Timeout::Error
    flunk "#{what} did not arrive within 30s"
  end

  # Polls a condition for up to 10s; a VBA handler runs inside Excel's
  # own call and is normally done before the put returns, but the bound
  # keeps a slow Wine from turning that into a flaky failure.
  def settled(seconds: 10)
    deadline = Process.clock_gettime(Process::CLOCK_MONOTONIC) + seconds
    loop do
      return true if yield
      return false if Process.clock_gettime(Process::CLOCK_MONOTONIC) > deadline

      sleep 0.2
    end
  end

  def timed(label)
    started = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    yield
    ms = ((Process.clock_gettime(Process::CLOCK_MONOTONIC) - started) * 1000).round
    puts "[measure] #{label}: #{ms} ms total"
  end
end
