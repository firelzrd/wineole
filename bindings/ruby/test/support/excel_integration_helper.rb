require 'timeout'
require_relative '../../lib/wineole'

# Shared spawn/teardown plumbing for the real-Excel integration tests.
#
# Extracted out of wineole_integration_test.rb so that file and
# wineole/msoffice_integration_test.rb share one copy of the bridge
# bring-up/tear-down dance instead of each maintaining a slightly
# different one. Getting one copy of this subtly wrong is how a stray
# EXCEL.EXE survives a test run -- see the comments below, carried over
# from the original.
#
# A fixed port, deliberately: a random one can never reuse an
# already-running bridge (standalone-started, or left by a previous run),
# which is the whole point of the connect-or-spawn logic -- and it risks
# colliding with an unrelated service. With the connection-teardown handle
# release in place (server.rs) and the PID-scoped cleanup below, repeated
# runs on one port are clean.
module ExcelIntegrationHelper
  BRIDGE_EXE = WineOLE::Client.default_bridge_path
  PORT = 48042

  def setup
    @spawned_pids = []
    @lockfile = WineOLE::Client.default_lockfile(PORT)
    # Excel is started by COM activation, not by us, so it has no PID we can
    # capture at spawn time. Snapshot what was already running instead, so
    # teardown can clean up only what this run caused -- never someone
    # else's Excel, possibly with unsaved work, as the old
    # `pkill -f EXCEL.EXE` did.
    @pre_existing_excel_pids = excel_pids
  end

  def teardown
    # Kill only what this run started, by PID. The old `pkill -f EXCEL.EXE`
    # killed every Excel on the machine -- including unrelated ones with
    # unsaved work -- and never killed the bridge at all, leaking one
    # bridge process per run for the full 30-minute idle timeout.
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

  # Split into two layers: with_bridge brings up the bridge and yields a
  # bare client (no Excel instance), and with_excel builds on it, adding
  # only the Excel-specific bring-up/teardown. This lets a test that needs
  # only a client (e.g. to hand to WineOLE::MSOffice::Excel.create/.run)
  # reuse the spawn/teardown logic without with_excel's own
  # client.create call defeating its "nothing exists yet" precondition.
  def with_bridge
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

    begin
      yield client
    ensure
      WineOLE.close
    end
  end

  def with_excel
    with_bridge do |client|
      # Inside the begin/ensure (via with_bridge's own ensure) so a raise
      # here still runs WineOLE.close -- a create failure must not skip
      # the graceful close and module-state reset.
      xl = client.create('Excel.Application')
      begin
        yield xl
      ensure
        quit_bounded(xl)
      end
    end
  end

  # `Quit` in a teardown is a Client#call on a connection the test body may
  # have wedged, and an unbounded one turns a failing test into a hanging
  # suite -- no output, no result, and every later test taken with it.
  # Measured on the events test that exists to catch callbacks running on the
  # reader thread: under that mutation the run produced nothing for 200s and
  # had to be killed, against 1.6s for a correct answer. With a clock on it
  # the same mutation is an ordinary failure.
  #
  # Timeout::Error is a StandardError, so the one rescue covers both a bridge
  # that refused the Quit and a bridge that never answered.
  def quit_bounded(xl, seconds: 20)
    Timeout.timeout(seconds) { xl.Quit }
  rescue StandardError
    nil
  end

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
    # -x matches the process name exactly; -f would match any command line
    # that merely mentions EXCEL.EXE (a diagnostic pgrep, a filesystem
    # search for the file), which silently skipped a test that requires
    # "no Excel running" and, worse, handed teardown an unrelated process
    # to kill.
    `pgrep -x EXCEL.EXE`.split.map(&:to_i)
  rescue StandardError
    []
  end

  # Checking for a leftover Excel process immediately after a run is racy:
  # the process can be mid-teardown (Wine's own COM release/exit sequence)
  # and gone moments later, so poll with a short timeout rather than
  # sampling once.
  def excel_gone?(timeout:)
    deadline = Time.now + timeout
    loop do
      return true if (excel_pids - @pre_existing_excel_pids).empty?
      return false if Time.now > deadline

      sleep 0.5
    end
  end
end
