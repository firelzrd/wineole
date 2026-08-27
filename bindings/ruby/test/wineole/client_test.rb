require 'minitest/autorun'
require 'socket'
require 'json'
require 'timeout'
require 'weakref'
require_relative '../../lib/wineole/client'

class ClientTest < Minitest::Test
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

    client = WineOLE::Client.new(TCPSocket.new('127.0.0.1', port))
    server_thread.join

    weak = WeakRef.new(client)
    client = nil
    GC.start

    refute weak.weakref_alive?, 'the Client must actually be collectible (finalizer must not capture self)'

    result = begin
      Timeout.timeout(2) { accepted.gets }
    rescue Timeout::Error
      :timed_out
    end
    assert_nil result, 'the peer must see EOF once the finalizer closed the socket (or the wait timed out)'

    accepted.close
    server.close
  end
end
