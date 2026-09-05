require 'minitest/autorun'
require 'socket'
require 'json'
require 'timeout'
require 'stringio'
require_relative '../../lib/wineole/client'
require_relative '../support/garbage_collection_helper'

class ClientEventsTest < Minitest::Test
  include GarbageCollectionHelper

  # A fake bridge on a real socket pair. Real enough to exercise the reader
  # thread and the framing; no COM, no Excel.
  def with_fake_bridge
    server, client_side = UNIXSocket.pair
    client = WineOLE::Client.new(client_side)
    yield client, server
  ensure
    # Bounded, because teardown runs inside the test method: a client whose
    # close cannot complete -- exactly what a lock held across a round trip
    # produces -- would otherwise hang here and swallow the verdict the test
    # had already reached.
    begin
      Timeout.timeout(5) { client&.close }
    rescue StandardError
      nil
    end
    server&.close rescue nil
  end

  # A sink that raises is reported on stderr; the tests that provoke one do
  # not need to see it.
  def quiet
    original = $stderr
    $stderr = StringIO.new
    yield
  ensure
    $stderr = original
  end

  def test_an_event_frame_does_not_break_a_pending_response
    with_fake_bridge do |client, server|
      got = nil
      caller = Thread.new { got = client.call('invoke', {handle: 1}) }

      request = JSON.parse(server.gets)
      # The event arrives FIRST, before the response. The old client read the
      # next line as its own response and raised "id mismatch"; that is the
      # regression this pins.
      server.write(JSON.generate({event: 'SheetChange', handle: 1, seq: 5, args: nil}) + "\n")
      server.write(JSON.generate({id: request['id'], result: 42}) + "\n")

      caller.join(5)
      assert_equal 42, got
    end
  end

  def test_events_reach_the_registered_consumer
    with_fake_bridge do |client, server|
      seen = Queue.new
      client.on_event { |frame| seen << frame }
      server.write(JSON.generate({event: 'Click', handle: 3, seq: 9, args: nil}) + "\n")

      frame = Timeout.timeout(5) { seen.pop }
      # The whole frame, not a field at a time: what this client owes a
      # consumer is the parsed frame exactly as it arrived, with nothing
      # dropped, renamed or filled in. (An earlier `assert_nil frame['args']`
      # here named a contract enforced on the Rust side -- true of any
      # pass-through, and of an implementation doing nothing at all.)
      assert_equal({'event' => 'Click', 'handle' => 3, 'seq' => 9, 'args' => nil}, frame,
                   'the frame must reach the consumer exactly as it arrived')
    end
  end

  # One connection can carry several objects with events on them (an
  # Application and a Workbook, say), each with its own consumer. A
  # registration that REPLACED the previous one would silently switch the
  # earlier objects' events off.
  def test_on_event_appends_rather_than_replacing_the_consumer
    with_fake_bridge do |client, server|
      first = Queue.new
      second = Queue.new
      client.on_event { |frame| first << frame }
      client.on_event { |frame| second << frame }

      server.write(JSON.generate({event: 'Click', handle: 3, seq: 9, args: nil}) + "\n")

      assert_equal 'Click', Timeout.timeout(5) { first.pop }['event'],
                   'the consumer registered first must still receive events'
      assert_equal 'Click', Timeout.timeout(5) { second.pop }['event'],
                   'the consumer registered second must receive events too'
    end
  end

  # The way back out, and why there has to be one: a consumer registered for
  # the life of the connection cannot be dismantled. WineOLE::Events takes
  # its sink off when its last callback goes, and without this the reader
  # would walk an entry for every object that ever had a callback on it --
  # each entry holding that object, and its dispatcher thread, alive until
  # the connection closed.
  def test_off_event_stops_one_consumer_and_leaves_the_others
    with_fake_bridge do |client, server|
      dropped = Queue.new
      kept = Queue.new
      # Registered FIRST, so it is called first while it is registered: if it
      # were still there after the off, its frame would be in the queue
      # before the one this test waits for.
      going = proc { |frame| dropped << frame }
      client.on_event(&going)
      client.on_event { |frame| kept << frame }

      server.write(JSON.generate({event: 'Click', handle: 3, seq: 9, args: nil}) + "\n")
      assert_equal 'Click', Timeout.timeout(5) { dropped.pop }['event'],
                   'it must receive before the off, or what follows proves nothing'
      Timeout.timeout(5) { kept.pop }

      client.off_event(going)
      server.write(JSON.generate({event: 'Other', handle: 3, seq: 10, args: nil}) + "\n")

      assert_equal 'Other', Timeout.timeout(5) { kept.pop }['event'],
                   'the consumer that stayed must go on receiving'
      assert_empty dropped, 'the removed consumer must receive nothing further'
      assert_equal 1, client.instance_variable_get(:@event_sinks).length,
                   'and it must not be left holding the connection'
    end
  end

  # Two callers in flight at once. The old client held a mutex across the
  # whole round trip, so a call made from inside an event callback could not
  # even be sent until the outer call returned.
  def test_two_threads_can_have_requests_in_flight_at_once
    with_fake_bridge do |client, server|
      a = Thread.new { client.call('invoke', {n: 1}) }
      first = JSON.parse(server.gets)
      b = Thread.new { client.call('invoke', {n: 2}) }
      # A plain `gets` here would HANG rather than fail if the second request
      # never reached the wire, so the wait is bounded and its expiry is the
      # assertion below.
      second = begin
        Timeout.timeout(5) { JSON.parse(server.gets) }
      rescue Timeout::Error
        nil
      end

      refute_nil second, 'the second request must reach the wire while the first is unanswered'
      # Answer them out of order, to prove the routing is by id and not by
      # arrival order.
      server.write(JSON.generate({id: second['id'], result: 'second'}) + "\n")
      server.write(JSON.generate({id: first['id'], result: 'first'}) + "\n")

      # `a.value` on its own blocks forever on a caller that is never answered
      # -- the routing bug this test names -- so each wait is bounded and the
      # expired join is what fails.
      assert a.join(5), 'the first caller must be answered'
      assert b.join(5), 'the second caller must be answered'
      assert_equal 'first', a.value
      assert_equal 'second', b.value
    end
  end

  def test_a_closed_connection_wakes_every_waiter
    with_fake_bridge do |client, server|
      waiters = 3.times.map { Thread.new { client.call('invoke', {}) rescue $! } }
      3.times { server.gets }
      server.close

      # `t.value` on its own would block forever on a waiter that was never
      # woken -- exactly the bug -- so an expired join contributes nil and the
      # count below is what fails.
      results = waiters.map { |t| t.join(5) ? t.value : nil }
      assert_equal 3, results.count { |r| r.is_a?(WineOLE::ProtocolError) },
                   'every waiter must be woken on EOF, or it waits forever'
    end
  end

  # Every consumer on a connection shares the one reader thread, so an
  # exception raised out of one of them would otherwise end the read loop:
  # the other consumers would go silent and every later call would fail with
  # "connection closed".
  def test_a_raising_event_consumer_does_not_take_the_connection_down
    # `quiet` wraps the whole fixture, not just the writes: the broken sink is
    # called once more, with nil, when the fixture closes the connection.
    quiet do
      with_fake_bridge do |client, server|
        seen = Queue.new
        client.on_event { |_frame| raise 'this consumer is broken' }
        client.on_event { |frame| seen << frame }

        server.write(JSON.generate({event: 'Click', handle: 3, seq: 1, args: nil}) + "\n")
        assert_equal 'Click', Timeout.timeout(5) { seen.pop }['event'],
                     'a consumer registered after a broken one must still see events'

        # ...and the connection must still carry a round trip.
        pending = Thread.new { client.call('invoke', {}) }
        request = Timeout.timeout(5) { JSON.parse(server.gets) }
        server.write(JSON.generate({id: request['id'], result: 'alive'}) + "\n")
        # Bounded for the same reason: a reader killed by the broken sink
        # would leave this caller waiting forever rather than failing.
        assert pending.join(5), 'the connection must still answer a call'
        assert_equal 'alive', pending.value
      end
    end
  end

  # Task 6's Events parks a dispatcher thread on a queue that only the reader
  # fills. Without an end-of-stream hand-off that thread blocks on an empty
  # queue for the life of the process, so every consumer -- not just the
  # waiters -- has to be told the stream is over.
  def test_a_closed_connection_tells_every_event_consumer_the_stream_ended
    with_fake_bridge do |client, server|
      first = Queue.new
      second = Queue.new
      client.on_event { |frame| first << [frame] }
      client.on_event { |frame| second << [frame] }

      server.close

      assert_nil Timeout.timeout(5) { first.pop }.first,
                 'the first consumer must be handed nil at EOF'
      assert_nil Timeout.timeout(5) { second.pop }.first,
                 'the second consumer must be handed nil at EOF'
    end
  end

  # The realistic path: the bridge dies, and only then does user code attach a
  # handler to a proxy it already had. `fail_all_waiters` dispatched its `nil`
  # once, to whoever was registered at that moment, so a consumer arriving
  # afterwards was never told anything -- and its dispatcher thread would park
  # on an empty queue for the life of the process.
  def test_a_consumer_registered_after_the_stream_ended_is_told_so_at_once
    with_fake_bridge do |client, server|
      ended = Queue.new
      client.on_event { |frame| ended << [frame] }
      server.close
      assert_nil Timeout.timeout(5) { ended.pop }.first, 'the stream must have ended'

      late = Queue.new
      client.on_event { |frame| late << [frame] }

      assert_nil Timeout.timeout(5) { late.pop }.first,
                 'a consumer registered after the stream ended must be handed nil immediately'
    end
  end

  # The deadlock this whole design exists to prevent. Holding the mailbox
  # mutex across the dispatch leaves every other test in this file passing,
  # and yet a call issued from inside a consumer cannot even reach the wire:
  # the reader holds the lock that guards the socket write, on the very thread
  # the consumer runs on.
  def test_a_call_issued_from_inside_an_event_consumer_reaches_the_wire
    with_fake_bridge do |client, server|
      client.on_event do |frame|
        next if frame.nil?

        begin
          # Bounded: the reader thread is inside this block, so nobody is left
          # to route the response back and this call can never complete. That
          # is the documented contract (hand off, do not block) -- what is
          # under test is only that the request got out.
          Timeout.timeout(2) { client.call('invoke', {from: 'a consumer'}) }
        rescue StandardError
          nil
        end
      end

      server.write(JSON.generate({event: 'Click', handle: 3, seq: 1, args: nil}) + "\n")

      request = begin
        Timeout.timeout(5) { JSON.parse(server.gets) }
      rescue Timeout::Error
        nil
      end

      refute_nil request, 'a call made from inside an event consumer must reach the wire'
      assert_equal 'invoke', request['method']
    end
  end

  # Frames WITH an id are answers to a caller and nobody else's business.
  # Dispatching them to the sinks as well leaves every other test here
  # passing, while a consumer that reads `frame['handle']` would raise once
  # per RPC -- silently, since a raising sink is only warned about.
  def test_only_frames_without_an_id_reach_the_event_consumers
    with_fake_bridge do |client, server|
      seen = Queue.new
      client.on_event { |frame| seen << [frame] }

      pending = Thread.new { client.call('invoke', {}) }
      request = Timeout.timeout(5) { JSON.parse(server.gets) }
      # The response FIRST, then the event: whatever reaches the consumer
      # first is what the reader considers dispatchable.
      server.write(JSON.generate({id: request['id'], result: 'answered'}) + "\n")
      server.write(JSON.generate({event: 'Click', handle: 3, seq: 1, args: nil}) + "\n")

      assert pending.join(5), 'the caller must be answered'
      assert_equal 'answered', pending.value

      frame = Timeout.timeout(5) { seen.pop }.first
      assert_equal 'Click', frame['event'],
                   'a response frame must go to its waiter only, never to the event consumers'
      assert seen.empty?, 'nothing but the event frame may reach a consumer'
    end
  end

  # `null`, `123` and `[]` are all valid JSON and none of them is a frame.
  # An unparseable line was already skipped; a parseable non-object used to
  # raise NoMethodError out of `frame.key?`, which is not in the read loop's
  # rescue list -- so one such line killed the reader, and with it every
  # consumer and every later call on the connection.
  def test_a_valid_json_line_that_is_not_an_object_does_not_kill_the_reader
    with_fake_bridge do |client, server|
      seen = Queue.new
      client.on_event { |frame| seen << [frame] }

      server.write("null\n123\n[]\n\"a string\"\n")
      server.write(JSON.generate({event: 'Click', handle: 3, seq: 1, args: nil}) + "\n")

      frame = Timeout.timeout(5) { seen.pop }.first
      refute_nil frame,
                 'a parseable non-object line killed the reader: the consumer was handed end-of-stream'
      assert_equal 'Click', frame['event'], 'a parseable non-object line must be skipped, not fatal'

      # ...and the connection must still carry a round trip.
      pending = Thread.new { client.call('invoke', {}) }
      request = Timeout.timeout(5) { JSON.parse(server.gets) }
      server.write(JSON.generate({id: request['id'], result: 'alive'}) + "\n")
      assert pending.join(5), 'the reader must have survived the non-object lines'
      assert_equal 'alive', pending.value
    end
  end

  # A consumer runs on the reader thread, so `close` called from inside one is
  # that thread joining itself: ThreadError, raised out of the one method
  # whose job is to shut the connection down cleanly.
  def test_close_called_from_inside_an_event_consumer_does_not_raise
    with_fake_bridge do |client, server|
      outcome = Queue.new
      client.on_event do |frame|
        next if frame.nil?

        begin
          client.close
          outcome << :closed
        rescue StandardError => e
          outcome << e
        end
      end

      server.write(JSON.generate({event: 'Click', handle: 3, seq: 1, args: nil}) + "\n")

      result = Timeout.timeout(5) { outcome.pop }
      assert_equal :closed, result,
                   "close from inside a consumer must not raise (got #{result.inspect})"
    end
  end

  # The reader reaches its sinks through weak references, so something has to
  # hold them: the Client does (Client#on_event). Drop that and a single
  # `client.on_event { }` keeps working only until the next collection, after
  # which the connection silently stops delivering events.
  def test_a_registered_consumer_survives_a_garbage_collection
    with_fake_bridge do |client, server|
      seen = Queue.new
      client.on_event { |frame| seen << [frame] }

      collect_garbage

      server.write(JSON.generate({event: 'Click', handle: 3, seq: 1, args: nil}) + "\n")
      frame = Timeout.timeout(5) { seen.pop }.first
      refute_nil frame, 'the connection ended instead of delivering the event'
      assert_equal 'Click', frame['event'],
                   'a sink must live as long as the client it was registered on'
    end
  end

  # "The reader is dead but the socket is still open" -- the bridge stopped
  # answering, or half-closed. Without the closed check in `request` the write
  # succeeds, no reader is left to route an answer, and the caller waits on
  # its slot for the life of the process.
  def test_a_call_after_the_stream_ended_raises_rather_than_waiting_forever
    with_fake_bridge do |client, server|
      ended = Queue.new
      client.on_event { |frame| ended << [frame] }
      # Half-close: the client sees EOF and its reader finishes, while its own
      # socket stays perfectly writable.
      server.close_write
      assert_nil Timeout.timeout(5) { ended.pop }.first, 'the reader must have finished'

      error = nil
      caller = Thread.new do
        begin
          client.call('invoke', {})
        rescue StandardError => e
          error = e
        end
      end

      assert caller.join(5), 'a call with no reader left must raise, not wait forever'
      assert_instance_of WineOLE::ProtocolError, error
      assert_equal 'connection closed', error.message
    end
  end

  # A socket stand-in that answers one request and then ends the stream by
  # exception, with no wait in between: the window in which the reader has
  # filled a waiter and woken its caller, but the caller has not yet
  # re-acquired the mailbox mutex to delete itself from the table.
  class AnswerThenBreak
    def initialize
      @requests = Queue.new
      @answered = false
    end

    def gets
      raise IOError, 'the stream ends here, with nothing to wait for' if @answered

      raw = @requests.pop
      return nil if raw.nil?

      @answered = true
      JSON.generate({id: JSON.parse(raw)['id'], result: 'answered'}) + "\n"
    end

    def write(data)
      @requests << data
      data.bytesize
    end

    def close
      @requests << nil
    end
  end

  # The sweep at EOF must not overwrite an answer that was already delivered.
  # Last-write-wins in `Waiter#fill` turns a request the bridge really
  # answered into `ProtocolError: connection closed`.
  def test_the_end_of_the_stream_does_not_overwrite_a_delivered_response
    client = WineOLE::Client.new(AnswerThenBreak.new)
    result = nil
    error = nil

    caller = Thread.new do
      begin
        result = client.call('invoke', {})
      rescue StandardError => e
        error = e
      end
    end

    assert caller.join(5), 'the caller must not be left waiting'
    assert_nil error,
               "a request the bridge answered must not be reported as #{error.class}: #{error&.message}"
    assert_equal 'answered', result
  ensure
    begin
      Timeout.timeout(5) { client&.close }
    rescue StandardError
      nil
    end
  end
end
