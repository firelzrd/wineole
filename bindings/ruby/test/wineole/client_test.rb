require 'minitest/autorun'
require 'socket'
require 'json'
require 'timeout'
require 'weakref'
require_relative '../../lib/wineole/client'
require_relative '../support/garbage_collection_helper'

class ClientTest < Minitest::Test
  include GarbageCollectionHelper

  def test_call_returns_result_on_success
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'], result: {'pong' => true}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    assert_equal({'pong' => true}, client.call('ping', {}))

    server_thread.join
    server.close
  end

  def test_call_raises_remote_error_on_error_response
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'], error: {'class' => 'WIN32OLERuntimeError', 'message' => 'boom'}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    error = assert_raises(WineOLE::RemoteError) { client.call('invoke', {}) }
    assert_equal 'WIN32OLERuntimeError: boom', error.message

    server_thread.join
    server.close
  end

  # The end of the stream is handed to a waiter as a sentinel object, not as a
  # fabricated error frame, so nothing the bridge can say is mistaken for it.
  # While it was a frame, a bridge reporting a WineOLE::ProtocolError of its
  # own had it rewritten into a local error -- this client blaming itself for
  # something the bridge said.
  def test_an_error_the_bridge_reports_stays_a_remote_error_whatever_its_class
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'],
                                error: {'class' => 'WineOLE::ProtocolError',
                                        'message' => 'the bridge itself said this'}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    error = assert_raises(WineOLE::RemoteError) { client.call('invoke', {}) }
    assert_equal 'WineOLE::ProtocolError: the bridge itself said this', error.message

    server_thread.join
    server.close
  end

  # The one remote error class this client resolves to its own local class
  # instead of wrapping in a generic RemoteError -- so a caller can rescue
  # "this instance is closing" directly rather than inspecting
  # RemoteError#remote_class.
  def test_call_raises_instance_closing_error_when_the_bridge_reports_it
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'],
                                error: {'class' => 'WineOLE::InstanceClosingError',
                                        'message' => 'the instance is closing'}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    error = assert_raises(WineOLE::InstanceClosingError) { client.call('invoke', {}) }
    assert_equal 'the instance is closing', error.message

    server_thread.join
    server.close
  end

  def test_call_sends_incrementing_ids
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    received_ids = []
    server_thread = Thread.new do
      conn = server.accept
      2.times do
        req = JSON.parse(conn.gets)
        received_ids << req['id']
        conn.write(JSON.generate({id: req['id'], result: nil}) + "\n")
      end
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    client.call('ping', {})
    client.call('ping', {})

    server_thread.join
    server.close
    assert_equal [1, 2], received_ids
  end

  def test_create_delegates_to_proxy_create
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'], result: {'$ole_ref' => 7}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    proxy = client.create('Excel.Application')

    assert_instance_of WineOLE::Proxy, proxy
    assert_equal 7, proxy.ole_handle

    server_thread.join
    server.close
  end

  def test_connect_delegates_to_proxy_connect
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'], result: {'$ole_ref' => 9}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    proxy = client.connect('Excel.Application')

    assert_instance_of WineOLE::Proxy, proxy
    assert_equal 9, proxy.ole_handle

    server_thread.join
    server.close
  end

  def test_connect_or_create_delegates_to_proxy_connect_or_create
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    server_thread = Thread.new do
      conn = server.accept
      req = JSON.parse(conn.gets)
      conn.write(JSON.generate({id: req['id'], result: {'$ole_ref' => 11, 'created' => false}}) + "\n")
      conn.close
    end

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    proxy = client.connect_or_create('Excel.Application')

    assert_instance_of WineOLE::Proxy, proxy
    assert_equal 11, proxy.ole_handle
    assert_equal false, proxy.ole_created?

    server_thread.join
    server.close
  end

  def test_close_does_not_raise_when_called_twice
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    accepted = nil
    server_thread = Thread.new { accepted = server.accept }

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    server_thread.join

    client.close
    client.close

    accepted.close
    server.close
  end

  def test_finalizer_closes_the_socket_without_capturing_self
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    accepted = nil
    server_thread = Thread.new { accepted = server.accept }

    # Built on a thread that is dead before the question is asked, so this
    # thread's stack never held a pointer to the Client at all. `client = nil`
    # here would not do it: it does not unbuild the stack slot the collector
    # has already seen. See GarbageCollectionHelper.
    weak = weak_ref_to { WineOLE::Client.new(TCPSocket.new('127.0.0.1', port)) }
    server_thread.join

    assert_collected weak, 'the Client must actually be collectible (finalizer must not capture self)'

    result = begin
      Timeout.timeout(2) { accepted.gets }
    rescue Timeout::Error
      :timed_out
    end
    assert_nil result, 'the peer must see EOF once the finalizer closed the socket (or the wait timed out)'

    accepted.close
    server.close
  end

  # The shape Task 6's `Events` has: an object that holds the client and
  # registers an event consumer closing over itself. Every such registration
  # used to pin the Client through the running reader thread -- the identical
  # leak the Mailbox was extracted to fix, walked back in through the sink
  # list. The finalizer test above registers no consumer, so it never saw it.
  class EventConsumer
    def initialize(client)
      @client = client
      @frames = Queue.new
      client.on_event { |frame| @frames << frame }
    end
  end

  def test_finalizer_still_runs_when_an_event_consumer_closes_over_the_client
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    accepted = nil
    server_thread = Thread.new { accepted = server.accept }

    # Both the Client and the consumer that closes over it are built on a
    # thread that is dead before the question is asked, so neither leaves a
    # pointer anywhere on this thread's stack. The consumer is dropped by that
    # thread ending, not by assigning nil over it -- assigning nil does not
    # unbuild a slot the collector has already seen.
    weak = weak_ref_to do
      client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
      EventConsumer.new(client)
      client
    end
    server_thread.join

    assert_collected weak,
      'a Client with an event consumer registered on it must still be collectible'

    result = begin
      Timeout.timeout(2) { accepted.gets }
    rescue Timeout::Error
      :timed_out
    end
    assert_nil result, 'the peer must see EOF once the finalizer closed the socket (or the wait timed out)'

    accepted.close
    server.close
  end

  def test_loopback_is_true_for_a_connection_to_127_0_0_1
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    begin
      socket = TCPSocket.new('127.0.0.1', port)
      begin
        assert WineOLE::Client.new(socket).loopback?,
          'a connection to 127.0.0.1 must count as loopback'
      ensure
        socket.close
      end
    ensure
      server.close
    end
  end

  def test_loopback_matches_the_bridge_s_own_definition
    # The bridge decides whether a token is required with
    # IpAddr::is_loopback(), which is true for all of 127.0.0.0/8 and ::1 --
    # not just the literal 127.0.0.1. Path conversion keys off the same
    # notion of "local", so the two must not drift apart.
    #
    # This asserts only on IPAddr directly, not on Client -- it documents
    # that the stdlib's definition matches Rust's, but gives no coverage to
    # the plumbing in `loopback?` itself. See
    # test_loopback_is_true_for_a_connection_to_ipv6_loopback below for that.
    assert IPAddr.new('127.0.0.1').loopback?
    assert IPAddr.new('127.0.0.2').loopback?
    assert IPAddr.new('::1').loopback?
    refute IPAddr.new('192.168.1.50').loopback?
  end

  def test_loopback_is_true_for_a_connection_to_ipv6_loopback
    server = TCPServer.new('::1', 0)
    port = server.addr[1]
    begin
      socket = TCPSocket.new('::1', port)
      begin
        assert WineOLE::Client.new(socket).loopback?,
          'a connection to ::1 must count as loopback'
      ensure
        socket.close
      end
    ensure
      server.close
    end
  rescue Errno::EADDRNOTAVAIL, SocketError => e
    skip "IPv6 loopback is not available in this environment: #{e.class}: #{e.message}"
  end

  def test_loopback_is_false_once_the_peer_address_is_undeterminable
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    begin
      socket = TCPSocket.new('127.0.0.1', port)
      client = WineOLE::Client.new(socket)
      socket.close

      refute client.loopback?,
        'loopback? must fail closed (false), not raise, once the peer address cannot be determined'
    ensure
      server.close
    end
  end

  def test_try_connect_sets_tcp_nodelay_on_the_socket
    server = TCPServer.new('127.0.0.1', 0)
    port = server.addr[1]
    begin
      socket = WineOLE::Client.send(:try_connect, '127.0.0.1', port)
      refute_nil socket, 'try_connect should have connected to the listening server'
      begin
        # Insurance against the Nagle/delayed-ACK stall fixed on the bridge
        # side: requests already go out in a single write, so this changes
        # nothing today, but it keeps a future multi-write request path from
        # reintroducing a 40 ms per-RPC penalty.
        nodelay = socket.getsockopt(Socket::IPPROTO_TCP, Socket::TCP_NODELAY).int
        assert_equal 1, nodelay, 'try_connect must set TCP_NODELAY'
      ensure
        socket.close
      end
    ensure
      server.close
    end
  end
end
