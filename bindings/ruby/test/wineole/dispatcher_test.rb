require 'minitest/autorun'
require 'stringio'
require_relative '../../lib/wineole'

# The Dispatcher's $cleanup delivery: a $cleanup frame runs the client's
# on_cleanup closure ON THE DISPATCHER THREAD (the COM-safe context, like
# every other callback), then acks the bridge with release_event(seq) and
# wakes whoever is blocked in await_cleanup. The closure's own exception must
# not stop any of that.
class DispatcherTest < Minitest::Test
  # A fake Client that records what reached the bridge and hands its one sink
  # back so a test can push a frame in the way the reader would. Shaped like
  # the real Client for exactly the calls the Dispatcher makes on the cleanup
  # path: call / on_event / off_event / signal_cleanup_done.
  def cleanup_client(calls:, signalled:)
    client = Object.new
    client.define_singleton_method(:call) { |m, p| calls << [m, p]; nil }
    client.define_singleton_method(:on_event) { |&blk| @sink = blk }
    client.define_singleton_method(:off_event) { |_blk| @sink = nil }
    client.define_singleton_method(:signal_cleanup_done) { |seq| signalled << seq }
    client.define_singleton_method(:sink) { @sink }
    client
  end

  def test_cleanup_frame_runs_proc_on_dispatcher_thread_and_acks
    calls = []
    ran_on = []
    signalled = []
    client = cleanup_client(calls: calls, signalled: signalled)

    dispatcher = WineOLE::Dispatcher.new(client)
    dispatcher.register_cleanup(7, proc { ran_on << Thread.current })
    # Feed a $cleanup frame the way the reader would.
    client.sink.call({'event' => '$cleanup', 'handle' => 7, 'seq' => 99, 'args' => nil})
    dispatcher.drain_for_test

    assert_equal 1, ran_on.length, 'the closure runs exactly once'
    assert_equal dispatcher.thread_for_test, ran_on.first, 'on the dispatcher thread'
    assert_includes calls, ['release_event', {seq: 99}]
    assert_includes signalled, 99
  end

  def test_cleanup_still_acks_when_closure_raises
    calls = []
    signalled = []
    client = cleanup_client(calls: calls, signalled: signalled)

    dispatcher = WineOLE::Dispatcher.new(client)
    dispatcher.register_cleanup(7, proc { raise 'boom' })
    quiet do
      client.sink.call({'event' => '$cleanup', 'handle' => 7, 'seq' => 5, 'args' => nil})
      dispatcher.drain_for_test
    end

    assert_includes calls, ['release_event', {seq: 5}]
    assert_includes signalled, 5
  end

  # A raising closure warns to $stderr on purpose; keep it out of the run's
  # output, the same way events_test.rb does for its raising callbacks.
  def quiet
    original = $stderr
    $stderr = StringIO.new
    yield
  ensure
    $stderr = original
  end
end
