require 'minitest/autorun'
require 'socket'
require 'json'
require 'tmpdir'
require_relative '../../lib/wineole/client'

class ClientOpenTest < Minitest::Test
  # Accepts one connection and answers every JSON Lines request with a
  # generic success, recording what it was asked. Enough to observe the
  # handshake open performs.
  class FakeBridge
    attr_reader :requests

    def initialize(port = 0)
      @server = TCPServer.new('127.0.0.1', port)
      @requests = []
      @mutex = Mutex.new
      @thread = Thread.new { serve }
    end

    def port
      @server.addr[1]
    end

    def requests_seen
      @mutex.synchronize { @requests.dup }
    end

    def close
      @thread.kill
      @server.close
    rescue IOError
      nil
    end

    private

    def serve
      loop do
        conn = @server.accept
        Thread.new(conn) do |c|
          while (line = c.gets)
            req = JSON.parse(line)
            @mutex.synchronize { @requests << req }
            c.write(JSON.generate({id: req['id'], result: {'pong' => true}}) + "\n")
          end
        end
      end
    rescue IOError, Errno::EBADF
      nil
    end
  end

  def test_reuses_already_listening_server_without_spawning
    bridge = FakeBridge.new
    spawn_called = false
    spawner = ->(_port) { spawn_called = true }

    Dir.mktmpdir do |dir|
      client = WineOLE::Client.open(port: bridge.port, spawner: spawner, lockfile: File.join(dir, 'lock'))
      refute spawn_called
      client.close
    end
  ensure
    bridge&.close
  end

  def test_spawns_when_nothing_listening
    port = 47810 + rand(1000)
    spawned_port = nil
    bridge = nil
    spawner = lambda do |p|
      spawned_port = p
      bridge = FakeBridge.new(p)
    end

    Dir.mktmpdir do |dir|
      client = WineOLE::Client.open(port: port, spawner: spawner, lockfile: File.join(dir, 'lock'), timeout: 5)
      assert_equal port, spawned_port
      client.close
    end
  ensure
    bridge&.close
  end

  # --- design doc 7.1 step 1: "on success, ping to confirm protocol
  # compatibility, then use it" ---

  def test_pings_after_connecting_to_an_existing_bridge
    bridge = FakeBridge.new
    spawner = ->(_p) { flunk 'should not spawn' }

    Dir.mktmpdir do |dir|
      client = WineOLE::Client.open(port: bridge.port, spawner: spawner, lockfile: File.join(dir, 'lock'))
      client.close
    end

    seen = bridge.requests_seen
    assert_equal 1, seen.length
    assert_equal 'ping', seen.first['method']
    assert_equal({}, seen.first['params'])
  ensure
    bridge&.close
  end

  def test_pings_after_spawning_a_new_bridge
    port = 47930 + rand(60)
    bridge = nil
    spawner = ->(p) { bridge = FakeBridge.new(p) }

    Dir.mktmpdir do |dir|
      client = WineOLE::Client.open(port: port, spawner: spawner, lockfile: File.join(dir, 'lock'), timeout: 5)
      client.close
    end

    seen = bridge.requests_seen
    assert_equal 'ping', seen.first['method'],
      'the after-spawn path must handshake too, not only the reuse path'
  ensure
    bridge&.close
  end

  def test_token_is_sent_in_the_ping
    bridge = FakeBridge.new

    Dir.mktmpdir do |dir|
      client = WineOLE::Client.open(
        port: bridge.port,
        spawner: ->(_p) { flunk 'should not spawn' },
        lockfile: File.join(dir, 'lock'),
        token: 's3cret'
      )
      client.close
    end

    seen = bridge.requests_seen
    assert_equal 'ping', seen.first['method']
    assert_equal({'token' => 's3cret'}, seen.first['params'])
  ensure
    bridge&.close
  end

  def test_ping_failure_propagates
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    responder = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate(
        {id: req['id'], error: {'class' => 'WineOLE::AuthError', 'message' => 'invalid or missing token'}}
      ) + "\n")
    end

    Dir.mktmpdir do |dir|
      error = assert_raises(WineOLE::RemoteError) do
        WineOLE::Client.open(port: port, spawner: ->(_p) {}, lockfile: File.join(dir, 'lock'))
      end
      assert_equal 'WineOLE::AuthError', error.remote_class
    end
  ensure
    responder&.kill
    server.close
  end

  def test_socket_is_closed_when_handshake_fails
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    conn = nil
    responder = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate(
        {id: req['id'], error: {'class' => 'WineOLE::AuthError', 'message' => 'invalid or missing token'}}
      ) + "\n")
    end

    Dir.mktmpdir do |dir|
      assert_raises(WineOLE::RemoteError) do
        WineOLE::Client.open(port: port, spawner: ->(_p) {}, lockfile: File.join(dir, 'lock'))
      end
    end

    responder.join

    # The client-side socket must have been closed as part of unwinding the
    # failed handshake: from the server's end, reading past the error
    # response should now see EOF rather than the connection staying open
    # and leaking a file descriptor on the client side.
    assert_nil conn.gets, 'client must close its socket when the handshake fails'
  ensure
    responder&.kill
    conn&.close
    server.close
  end

  def test_default_lockfile_is_per_port
    assert_equal File.join(Dir.tmpdir, 'wineole-bridge.47800.lock'), WineOLE::Client.default_lockfile(47800)
    assert_equal File.join(Dir.tmpdir, 'wineole-bridge.48123.lock'), WineOLE::Client.default_lockfile(48123)
    refute_equal WineOLE::Client.default_lockfile(47800), WineOLE::Client.default_lockfile(47801),
      'two bridges on different ports must not contend for one lock'
  end

  def test_open_uses_the_per_port_default_lockfile
    port = 47990 + rand(10)
    lockfile = WineOLE::Client.default_lockfile(port)
    File.delete(lockfile) if File.exist?(lockfile)
    bridge = nil

    begin
      client = WineOLE::Client.open(
        port: port,
        spawner: ->(p) { bridge = FakeBridge.new(p) },
        timeout: 5
      )
      client.close
      assert File.exist?(lockfile), "expected #{lockfile} to have been used"
    ensure
      bridge&.close
      File.delete(lockfile) if File.exist?(lockfile)
    end
  end

  def test_raises_if_spawned_server_never_comes_up
    port = 47900 + rand(1000)
    spawner = ->(_p) {} # does nothing — nobody ever listens

    Dir.mktmpdir do |dir|
      assert_raises(WineOLE::Error) do
        WineOLE::Client.open(port: port, spawner: spawner, lockfile: File.join(dir, 'lock'), timeout: 1)
      end
    end
  end

  def test_default_bridge_path_resolves_relative_to_the_library_for_the_host_architecture
    path = WineOLE::Client.default_bridge_path
    expected_triple = case RbConfig::CONFIG['host_cpu']
                       when 'x86_64' then 'x86_64-pc-windows-gnu'
                       when /^i.86$/ then 'i686-pc-windows-gnu'
                       else nil
                       end
    skip "no prebuilt binary for host_cpu=#{RbConfig::CONFIG['host_cpu']}" if expected_triple.nil?

    assert_includes path, expected_triple
    assert_includes path, 'wineole-bridge.exe'
    assert File.exist?(path), "expected a prebuilt binary at #{path}"
  end

  def test_default_bridge_path_raises_for_an_unsupported_architecture
    error = assert_raises(WineOLE::Error) { WineOLE::Client.bridge_path_for_arch('sparc64') }
    assert_match(/sparc64/, error.message)
  end
end
