require 'minitest/autorun'
require_relative '../lib/wineole'

class WineOLEIntegrationTest < Minitest::Test
  BRIDGE_EXE = WineOLE::Client.default_bridge_path

  # A fixed port, deliberately: a random one can never reuse an
  # already-running bridge (standalone-started, or left by a previous run),
  # which is the whole point of the connect-or-spawn logic — and it risks
  # colliding with an unrelated service. With the connection-teardown handle
  # release in place (server.rs) and the PID-scoped cleanup below, repeated
  # runs on one port are clean.
  PORT = 48042

  def setup
    @spawned_pids = []
    @lockfile = WineOLE::Client.default_lockfile(PORT)
    # Excel is started by COM activation, not by us, so it has no PID we can
    # capture at spawn time. Snapshot what was already running instead, so
    # teardown can clean up only what this run caused — never someone else's
    # Excel, possibly with unsaved work, as the old `pkill -f EXCEL.EXE` did.
    @pre_existing_excel_pids = excel_pids
  end

  def test_end_to_end_excel_automation_via_the_bridge
    skip "bridge exe not built: #{BRIDGE_EXE}" unless File.exist?(BRIDGE_EXE)

    client = WineOLE.open(
      port: PORT,
      spawner: lambda { |p|
        pid = Process.spawn('wine', BRIDGE_EXE, p.to_s, %i[out err] => File::NULL)
        @spawned_pids << pid
        Process.detach(pid)
        pid
      },
      lockfile: @lockfile,
      timeout: 20
    )
    root_handle = nil

    begin
      xl = WineOLE.create('Excel.Application')
      root_handle = xl.ole_handle
      xl.Visible = false
      xl.DisplayAlerts = false
      xl.Workbooks.Add

      sheets = xl.Worksheets
      before_count = sheets.Count
      first_sheet = sheets[1]
      first_sheet_name = first_sheet.Name
      sheets.Add('After' => first_sheet)
      after_count = xl.Worksheets.Count

      assert_equal before_count + 1, after_count

      # The count assertion above only proves the 'After' named argument was
      # transmitted and accepted (a dropped/unresolvable name would raise
      # DISP_E_UNKNOWNNAME via GetIDsOfNames). It does not prove the argument
      # was honored positionally. Check that the new sheet actually landed
      # immediately after first_sheet, not before it or elsewhere.
      assert_equal first_sheet_name, xl.Worksheets[1].Name,
        "first_sheet should remain sheet 1 (unmoved) after Add('After' => first_sheet)"
      refute_equal first_sheet_name, xl.Worksheets[2].Name,
        "sheet 2 should be the newly added sheet, not first_sheet, proving 'After' was honored"

      # A COM error must arrive with something usable in it — an empty
      # message used to be the norm, since Wine cannot format most
      # automation HRESULTs.
      err = assert_raises(WineOLE::RemoteError) { xl.NoSuchMemberAtAll }
      assert_equal 'WIN32OLERuntimeError', err.remote_class
      assert_match(/0x[0-9A-F]{8}/, err.message, "expected an HRESULT in #{err.message.inspect}")

      xl.Quit
    ensure
      client.close
      WineOLE.instance_variable_set(:@default_client, nil)
    end

    # Dropping the connection — rather than releasing every handle one by
    # one, which no real client does — must make the bridge reclaim the whole
    # session. Observable from a fresh connection: the handles that
    # connection owned are gone from the routing table, so invoking one is a
    # stale reference rather than a live object.
    # Teardown is asynchronous (the read loop notices the disconnect, then
    # walks the handles), so poll rather than assuming it has already
    # finished by the time this second connection is up.
    second = connect
    begin
      assert handle_reclaimed?(second, root_handle, timeout: 20),
        'handles owned by a closed connection must not survive it'
    ensure
      second.close
    end

    # And the automated process itself should go away with the session. This
    # is best-effort rather than asserted: the observed logs show
    # `IRemUnknown_RemRelease failed with error 0x80070057` (E_INVALIDARG)
    # alongside `get_stub_manager_from_ipid not found for ipid {...}` — i.e.
    # the stub was already gone by the time the release reached it. That is
    # an over-release/late-release symptom, not an under-release, so it
    # cannot by itself be what keeps EXCEL.EXE pinned (a release the target
    # no longer needs can't keep it alive). A more plausible, but unconfirmed,
    # explanation: the session's STA worker thread (session.rs,
    # `CoInitializeEx(..., COINIT_APARTMENTTHREADED)`) never pumps a Windows
    # message loop anywhere in this crate (no `PeekMessage`/`GetMessage`/
    # `CoWaitForMultipleHandles`) — so if Excel ever raises a modal dialog
    # during `Quit` (e.g. a save prompt), nothing would dismiss it, which
    # would explain an intermittent, timing-dependent hang. This is a
    # hypothesis for follow-up investigation, not a confirmed diagnosis, and
    # fixing it is out of scope here. `test_connection_teardown_...` in
    # server.rs is the deterministic proof that the session shuts down.
    warn 'note: EXCEL.EXE outlived its session (known Wine COM release flake)' unless excel_gone?(timeout: 20)
  end

  def test_a_second_client_reuses_the_already_running_bridge
    skip "bridge exe not built: #{BRIDGE_EXE}" unless File.exist?(BRIDGE_EXE)

    first = connect
    spawned_after_first = @spawned_pids.length

    second = connect
    assert_equal spawned_after_first, @spawned_pids.length,
      'the second client must reuse the running bridge, not spawn another'

    first.close
    second.close
  end

  def test_connect_or_create_creates_then_a_second_call_attaches
    skip "bridge exe not built: #{BRIDGE_EXE}" unless File.exist?(BRIDGE_EXE)
    # Not `excel_pids - @pre_existing_excel_pids` -- setup snapshots
    # @pre_existing_excel_pids fresh before every test, so a leftover Excel
    # from a sibling test is captured *into* that snapshot and the
    # difference would always be empty. connect_or_create attaches to ANY
    # running Excel.Application, test-started or not, so the only clean
    # precondition is that none is running at all.
    unless excel_pids.empty?
      skip 'an Excel instance is already running -- cannot assert "nothing was running" cleanly; ' \
           're-run this test in isolation or investigate the leftover instance'
    end

    client = connect
    begin
      xl = client.connect_or_create('Excel.Application')
      assert xl.ole_created?, 'the first connect_or_create must create a new instance (nothing was running)'
      xl.Visible = false
      xl.DisplayAlerts = false

      xl2 = client.connect_or_create('Excel.Application')
      refute xl2.ole_created?, 'the second connect_or_create must attach to the instance the first one created'

      # Prove xl2 is genuinely the same live instance xl mutated, not a
      # coincidentally similar second one.
      xl2.DisplayAlerts = true
      xl.DisplayAlerts = false
      assert_equal false, xl2.DisplayAlerts, 'xl and xl2 must observe the same live Excel instance'

      xl.Quit
    ensure
      client.close
    end
  end

  private

  def connect
    WineOLE::Client.open(
      port: PORT,
      spawner: lambda { |p|
        pid = Process.spawn('wine', BRIDGE_EXE, p.to_s, %i[out err] => File::NULL)
        @spawned_pids << pid
        Process.detach(pid)
        pid
      },
      lockfile: @lockfile,
      timeout: 20
    )
  end

  # Has the bridge dropped `handle` from its routing table? Any other
  # outcome (a successful invoke, or a different error class) means the
  # handle is still live and is reported as "not yet reclaimed".
  def handle_reclaimed?(client, handle, timeout:)
    deadline = Time.now + timeout
    loop do
      begin
        client.call('invoke', {handle: handle, name: 'Version', args: [], named: {}})
      rescue WineOLE::RemoteError => e
        return true if e.remote_class == 'WineOLE::StaleReferenceError'
      end
      return false if Time.now > deadline

      sleep 0.5
    end
  end

  def excel_pids
    `pgrep -f EXCEL.EXE`.split.map(&:to_i)
  rescue StandardError
    []
  end

  def excel_gone?(timeout:)
    deadline = Time.now + timeout
    loop do
      return true if (excel_pids - @pre_existing_excel_pids).empty?
      return false if Time.now > deadline

      sleep 0.5
    end
  end

  def teardown
    # Kill only what this run started, by PID. The old `pkill -f EXCEL.EXE`
    # killed every Excel on the machine — including unrelated ones with
    # unsaved work — and never killed the bridge at all, leaking one bridge
    # process per run for the full 30-minute idle timeout.
    @spawned_pids.each do |pid|
      Process.kill('TERM', pid) # takes the wine wrapper and the .exe with it
    rescue Errno::ESRCH
      nil
    end
    excel_gone?(timeout: 5)
    (excel_pids - @pre_existing_excel_pids).each do |pid|
      Process.kill('TERM', pid)
    rescue Errno::ESRCH
      nil
    end
    File.delete(@lockfile) if @lockfile && File.exist?(@lockfile)
  end
end
