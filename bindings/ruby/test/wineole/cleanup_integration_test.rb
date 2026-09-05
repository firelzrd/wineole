require 'minitest/autorun'
require_relative '../../lib/wineole'
require_relative '../support/excel_integration_helper'

# The first end-to-end proof of the $cleanup client-closure path (Tasks
# 1-8) against a real Excel, driven through the low-level Client/Proxy API
# directly. MSOffice::Excel never uses `on_cleanup` (its CLEANUP_STEPS in
# excel.rb passes only `steps:`), so these two tests are the only place the
# closure delivery chain is exercised end-to-end rather than through mocks:
#
#   bridge emits `$cleanup` on last-user release (Task 5)
#     -> Dispatcher#run_cleanup runs the closure on the dispatcher thread,
#        the same COM-safe context every other callback runs on (Task 8)
#     -> the closure returns
#     -> `release_event` is sent and `Client#await_cleanup` unblocks (Task 7)
#     -> the bridge runs the steps (choice B: the closure is a PRELUDE to
#        shutdown, not a veto -- only `leave_open` can cancel it, edge 2
#        below).
#
# Shares ExcelIntegrationHelper's spawn/teardown plumbing with the other two
# integration files -- see that module for why getting the bring-up/tear-down
# dance wrong is exactly how a stray EXCEL.EXE survives a run.
class WineOLECleanupIntegrationTest < Minitest::Test
  include ExcelIntegrationHelper

  CLEANUP_STEPS = [['DisplayAlerts=', false], ['Quit']].freeze

  # Choice B: the closure is a prelude, not a veto. `ole_release` blocks
  # until the dispatcher has both run the closure and received the
  # `release_event` completion (Client#await_cleanup), so by the time it
  # returns the flag the closure sets must already be true. Whether the
  # EXCEL.EXE *process* has actually exited yet is a separate question --
  # the bridge issues the Quit step synchronously before `ole_release`
  # unblocks (measured at ~2.35ms in Task 6), but Wine's own process
  # teardown can lag well behind the COM call returning (the same flake
  # documented in msoffice_integration_test.rb's
  # test_run_quits_only_what_it_created), so this polls with a bounded
  # wait rather than asserting the PID is gone the instant control returns.
  def test_on_cleanup_closure_runs_then_steps_quit_excel
    with_bridge do |client|
      ran = false
      xl = client.create('Excel.Application',
        cleanup: {steps: CLEANUP_STEPS, on_cleanup: proc { ran = true }})
      xl.Visible = false
      xl.Workbooks.Add

      xl.ole_release # blocks until the $cleanup closure + steps complete

      assert ran, 'the on_cleanup closure must have run before ole_release returned'
      assert excel_gone?(timeout: 20),
        'the steps must still quit Excel after the closure runs (choice B)'
    end
  end

  # Edge 2: a closure that calls `ole_leave_open` on the root proxy revokes
  # the bridge's shutdown permission from inside the callback itself, so
  # the release that triggered the closure runs no steps at all -- the
  # instance survives both `ole_release` and the connection closing
  # (with_bridge's own ensure calls WineOLE.close right after this block).
  #
  # `root` is assigned to `xl` right after `create` returns, before
  # `ole_release` is ever called -- the closure only reads `root` once the
  # bridge actually asks for it, on the dispatcher thread, by which point
  # the local's value is set. Capturing `xl` directly instead would work
  # too (it does not depend on the reassignment); this mirrors the brief's
  # own phrasing of "bind the root proxy into the closure".
  def test_on_cleanup_closure_calling_leave_open_keeps_excel_open
    with_bridge do |client|
      root = nil
      xl = client.create('Excel.Application',
        cleanup: {steps: CLEANUP_STEPS, on_cleanup: proc { root.ole_leave_open }})
      root = xl
      xl.Visible = false
      xl.Workbooks.Add

      xl.ole_release

      # A grace period, not a formality: measured directly (3 runs) with
      # `on_cleanup` mutated to a no-op that never calls `ole_leave_open` --
      # `ole_release` still returns in ~2ms (the bridge's Quit step is
      # issued synchronously before it unblocks), but the EXCEL.EXE PID
      # lingered for close to 1s past that before actually exiting, on
      # every run. Checking `excel_pids` the instant `ole_release` returns
      # would therefore pass this assertion even with a broken
      # `ole_leave_open` that never revoked shutdown -- a false positive
      # from a process that is merely still tearing down, not one this
      # closure actually kept alive. Sleeping past that window first is
      # what makes "still running" mean the steps never ran at all, per
      # `confirm_cleanup` returning false (registry.rs), rather than "ran,
      # but hasn't finished exiting yet". 3s is comfortably above the ~1s
      # observed, while staying well under the separate, rare Wine
      # COM-release flake documented in msoffice_integration_test.rb
      # (`test_run_quits_only_what_it_created`) where a *genuine* Quit's
      # process teardown was seen to take over 20s -- not a concern here,
      # because the correct `ole_leave_open` path never issues Quit at all.
      sleep 3

      refute_empty excel_pids - @pre_existing_excel_pids,
        'a closure calling ole_leave_open must keep Excel running past ole_release'
    end
    # Left running deliberately -- teardown (ExcelIntegrationHelper) kills
    # any PID not in @pre_existing_excel_pids, so this does not leak.
  end
end
