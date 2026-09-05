require 'minitest/autorun'
require 'stringio'
require 'socket'
require 'json'
require 'weakref'
require_relative '../../lib/wineole'
require_relative '../support/garbage_collection_helper'

class EventsTest < Minitest::Test
  include GarbageCollectionHelper

  # Records what reached the bridge, and lets a test push frames in as if
  # they had arrived on the socket.
  class FakeClient
    attr_reader :calls, :sinks

    def initialize
      @calls = []
      @sinks = []
    end

    def call(method, params = {})
      @calls << [method, params]
      true
    end

    # One per connection, exactly as the real Client has one -- which is the
    # property most of the tests below are about. Lazily rather than in
    # `initialize` only because a fake needs no answer to the question the
    # real Client builds it eagerly for (two threads registering callbacks at
    # once); every test here reaches it from one thread first.
    def dispatcher = @dispatcher ||= WineOLE::Dispatcher.new(self)

    # Appends, mirroring the real Client -- a fake that replaced the sink
    # could not reach the two-objects-on-one-connection case at all.
    def on_event(&b) = @sinks << b
    # By identity, mirroring the real Client for the same reason it does.
    def off_event(b) = @sinks.reject! { |s| s.equal?(b) }
    def deliver(frame) = @sinks.dup.each { |s| s.call(frame) }
  end

  # A client whose `call` for one method blocks until the test lets it
  # through. That is what makes an interleaving deterministic rather than
  # hopeful: the thread under test is held INSIDE the wire call, in exactly
  # the window where it has decided something and not yet carried it out.
  class GatedClient < FakeClient
    def initialize(gated_method)
      super()
      @gated_method = gated_method
      @entered = Queue.new
      @release = Queue.new
    end

    def call(method, params = {})
      if method == @gated_method
        @entered << method
        @release.pop
      end
      super
    end

    def wait_until_inside(seconds = 5) = @entered.pop(timeout: seconds)
    def let_it_through = @release << :go
  end

  # The same idea for one call only: the FIRST release_event blocks, and
  # every one after it goes straight through. One-shot on purpose -- the
  # window this holds open is the dispatcher's idle hand-off, and everything
  # the test does after it has to be able to complete.
  class OnceReleaseGatedClient < FakeClient
    def initialize
      super()
      @entered = Queue.new
      @release = Queue.new
      @already_gated = false
    end

    def call(method, params = {})
      if method == 'release_event' && !@already_gated
        @already_gated = true
        @entered << params[:seq]
        @release.pop
      end
      super
    end

    def wait_until_inside(seconds = 5) = @entered.pop(timeout: seconds)
    def let_it_through = @release << :go
  end

  def events_for(client = FakeClient.new, handle: 1)
    [WineOLE::Events.new(client, handle), client]
  end

  def frame(event, seq, args: nil, handle: 1)
    {'event' => event, 'handle' => handle, 'seq' => seq, 'args' => args}
  end

  # A bounded `pop`, and the only deviation from the brief's text for these
  # tests: every assertion below keeps the value it was written with. The
  # dispatcher hands its result over a queue, and under the bug each of these
  # tests names -- a callback that never runs, an error that never reaches
  # on_error, a frame delivered to the wrong Events -- nothing is ever
  # pushed, so a bare `pop` would hang the suite forever instead of failing.
  # Timing out returns nil, which fails the same assertion in seconds with
  # the expected value still in the message.
  def waited(queue, seconds = 5)
    queue.pop(timeout: seconds)
  end

  # L1/L2 are derived: registering a callback is the ONLY thing the caller
  # does, and the subscribe must follow from it. If it did not, the callback
  # would be registered and the event would never arrive -- silently.
  def test_registering_a_callback_subscribes_on_the_bridge
    ev, client = events_for
    ev.on('SheetChange') { }
    assert_equal [['subscribe', {handle: 1, event: 'SheetChange', args: true}]], client.calls
  end

  # An object that is not an event source is refused by the bridge, and that
  # refusal arrives here as an exception out of `on`. If the callback stayed
  # registered anyway, the caller would be left holding a Subscription for an
  # event that can never arrive -- the one state this class is built to make
  # unreachable -- and a later `on` for the same event would not even try to
  # subscribe again, because a callback would already be on the list.
  def test_a_refused_subscribe_leaves_no_callback_behind
    client = FakeClient.new
    def client.call(method, params = {})
      super
      raise WineOLE::RemoteError.new('WIN32OLERuntimeError', 'not an event source') if method == 'subscribe'

      true
    end
    ev = WineOLE::Events.new(client, 1)

    assert_raises(WineOLE::RemoteError) { ev.on('Click') { } }
    assert_raises(WineOLE::RemoteError) { ev.on('Click') { } }
    assert_equal 2, client.calls.count { |m, _| m == 'subscribe' },
      'the refused registration must be gone, so the next `on` tries again rather ' \
      'than finding a callback already listed and subscribing to nothing'
  end

  def test_a_second_callback_for_the_same_event_does_not_subscribe_again
    ev, client = events_for
    ev.on('SheetChange') { }
    ev.on('SheetChange') { }
    assert_equal 1, client.calls.count { |m, _| m == 'subscribe' }
  end

  def test_removing_the_last_callback_unsubscribes
    ev, client = events_for
    a = ev.on('SheetChange') { }
    ev.on('SheetChange') { }
    ev.off(a)
    assert_equal 0, client.calls.count { |m, _| m == 'unsubscribe' }, 'one left, still subscribed'
    ev.off('SheetChange')
    assert_equal 1, client.calls.count { |m, _| m == 'unsubscribe' }
  end

  # The bridge holds one args flag per event, and `on` is the only place the
  # caller states what it wants. Measured on Excel before this was derived:
  # a second callback asking for arguments next to an `args: false` one was
  # handed nil, because the flag from the first registration was still
  # standing and nothing re-subscribed. The wire flag is the union of the
  # live callbacks, in both directions.
  def test_the_wire_flag_follows_what_the_callbacks_asked_for
    ev, client = events_for
    ev.on('Click', args: false) { }
    wants_args = ev.on('Click') { }
    assert_equal [
      ['subscribe', {handle: 1, event: 'Click', args: false}],
      ['subscribe', {handle: 1, event: 'Click', args: true}],
    ], client.calls, 'a callback that wants the arguments must re-subscribe for them'

    ev.off(wants_args)
    assert_equal ['subscribe', {handle: 1, event: 'Click', args: false}], client.calls.last,
      'and with it gone, stop paying for handles nobody asked for'
    assert_equal 0, client.calls.count { |m, _| m == 'unsubscribe' },
      'one callback is left, so the subscription itself stays'
  end

  # `off` removes the registration it was handed and no other. The same proc
  # registered twice is two registrations, and value equality cannot tell
  # them apart: measured on the Struct this used to be, s1 == s2 (without
  # being `equal?`), so `off(s1)` removed both, sent the unsubscribe, and s2
  # never fired again with nothing said about it.
  def test_off_removes_only_the_registration_it_was_given
    ev, client = events_for
    seen = Queue.new
    same_callback = proc { seen << :ran }
    first = ev.on('Click', &same_callback)
    second = ev.on('Click', &same_callback)
    refute_same first, second, 'two registrations, whatever they hold'

    ev.off(first)
    assert_equal 0, client.calls.count { |m, _| m == 'unsubscribe' },
      'the second registration is still there, so the subscription must stay'
    client.deliver(frame('Click', 1))
    assert_equal :ran, waited(seen), 'the registration that was not named must still fire'
  end

  # This project's own discipline: assert it fired BEFORE, then stopped.
  # "Nothing arrived" on its own is what a test that never worked also says.
  # Two events, in order, on the one dispatcher: if the removed callback
  # still ran, its :click reaches the queue first and the assertion says so.
  def test_off_stops_delivery
    ev, client = events_for
    seen = Queue.new
    click = ev.on('Click') { seen << :click }
    # A second event keeps the object armed, so what this test measures is
    # `off`, not the teardown that follows the last callback.
    ev.on('Other') { seen << :other }

    client.deliver(frame('Click', 1))
    assert_equal :click, waited(seen), 'it must fire before the off, or what follows proves nothing'

    ev.off(click)
    client.deliver(frame('Click', 2))
    client.deliver(frame('Other', 3))
    assert_equal :other, waited(seen), 'the removed callback must not have fired'
  end

  # `after.nil?` cannot tell "the last callback just went" from "there was
  # never one", and an unsubscribe for a subscription that does not exist is
  # a round trip that says nothing -- on an object the bridge may never have
  # advised at all.
  def test_off_for_a_name_that_was_never_registered_touches_nothing
    ev, client = events_for
    ev.off('NeverRegistered')

    assert_empty client.calls, 'there is no subscription to take down'
    assert_empty ev.registered_names_for_test, 'and a read must not leave a key behind'
    assert_empty client.sinks, 'nor put a consumer on the connection'
  end

  # The same rule from the other side: one connection's frames reach every
  # Events on it, so names nobody registered for arrive here all the time. A
  # table that grew a permanent key for each of them is a read that mutates.
  def test_an_event_nobody_registered_for_leaves_no_trace
    ev, client = events_for
    ev.on('Click') { }
    client.deliver(frame('Unwanted', 5))
    ev.drain_for_test

    assert_equal ['Click'], ev.registered_names_for_test
  end

  # `args: null` says the bridge minted NOTHING for this event -- it is sent
  # as null rather than left out precisely so the client can tell that from
  # "zero arguments" (protocol.rs). release_event is a synchronous round trip
  # from the dispatcher, so sending one anyway caps the event rate at one RTT
  # in exactly the high-frequency case `args: false` exists for.
  def test_an_event_that_minted_nothing_is_not_released
    ev, client = events_for
    ran = Queue.new
    ev.on('Click', args: false) { ran << :ran }
    client.deliver(frame('Click', 9))
    assert_equal :ran, waited(ran)
    ev.drain_for_test

    assert_equal 0, client.calls.count { |m, _| m == 'release_event' },
      'there is nothing to give back'
  end

  # Every unit frame above carries nil or [] for `args`, so the decode this
  # exercises -- the one that turns an event argument into something a
  # callback can use -- had no unit coverage at all: a build_args that never
  # wrapped a $ole_ref passed the whole file.
  def test_event_arguments_arrive_decoded_like_any_other_result
    ev, client = events_for
    got = Queue.new
    ev.on('Click') { |*args| got << args }

    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 5, 'args' => [
      {'$ole_ref' => 77},
      {'$type' => 'time', 'iso8601' => '2026-08-31T09:30:45'},
      'plain',
      42,
    ]})
    args = waited(got)

    assert_instance_of WineOLE::Proxy, args[0], 'an object argument must arrive callable'
    assert_equal 77, args[0].ole_handle
    assert_instance_of Time, args[1], 'and a date as a Time, exactly as a call result would'
    assert_equal 9, args[1].hour
    assert_equal 'plain', args[2]
    assert_equal 42, args[3]
  end

  # @subs is mutated under one lock and the bridge is told outside it, so two
  # threads can reach the wire in the opposite order to the one they decided
  # in. Measured on the code before @wire_mutex: wire order subscribe,
  # subscribe, unsubscribe -- with a callback still registered, i.e. a
  # callback whose event can never arrive.
  def test_a_subscribe_must_not_land_after_the_unsubscribe_that_follows_it
    client = GatedClient.new('unsubscribe')
    ev = WineOLE::Events.new(client, 1)
    first = ev.on('Click') { }

    remover = Thread.new { ev.off(first) }
    assert client.wait_until_inside, 'the remover must reach its wire call'
    adder = Thread.new { ev.on('Click') { } }
    # Under the fix the adder cannot even begin to decide until the remover
    # is done with the wire, so this join is expected to time out. It is not
    # asserted either way -- what matters is the state both orders end in.
    adder.join(1)
    client.let_it_through
    assert remover.join(5), 'the remover must finish'
    assert adder.join(5), 'the adder must finish'

    assert_equal ['Click'], ev.registered_names_for_test, 'a callback is registered'
    assert_equal 'subscribe', client.calls.last.first,
      'so the last thing the bridge heard must be a subscribe -- an unsubscribe here is a ' \
      'registered callback whose event can never arrive'
  end

  # The mirror: the bridge left advised with nothing to deliver to. A leaked
  # Advise, and every event it goes on raising minting handles nobody
  # releases.
  def test_an_unsubscribe_must_not_land_after_the_subscribe_that_follows_it
    client = GatedClient.new('subscribe')
    ev = WineOLE::Events.new(client, 1)

    adder = Thread.new { ev.on('Click') { } }
    assert client.wait_until_inside, 'the adder must reach its wire call'
    remover = Thread.new { ev.off('Click') }
    remover.join(1)
    client.let_it_through
    assert adder.join(5), 'the adder must finish'
    assert remover.join(5), 'the remover must finish'

    assert_empty ev.registered_names_for_test, 'no callback is registered'
    assert_equal 'unsubscribe', client.calls.last.first,
      'so the last thing the bridge heard must be an unsubscribe -- a subscribe here leaves ' \
      'the object advised with nobody to deliver to'
  end

  # A callback runs on the dispatcher, and the dispatcher is what `off`
  # takes down when the last callback goes. Registering and removing from
  # inside one must therefore work, and must not deadlock against the lock
  # the callback's own delivery took on the way in.
  def test_a_callback_can_register_and_remove_from_inside_the_dispatcher
    ev, client = events_for
    survived = Queue.new
    seen = Queue.new
    sub = nil
    sub = ev.on('Click') do
      # This order on purpose: the off empties the registry, which takes the
      # thread this callback is running on down with it, and the on puts it
      # straight back. One dispatcher must come out of that, not two and not
      # none.
      ev.off(sub)
      ev.on('Other') { seen << :other }
      survived << :survived
    end

    client.deliver(frame('Click', 1))
    assert_equal :survived, waited(survived), 'a callback calling off/on must not deadlock'
    client.deliver(frame('Other', 2))
    assert_equal :other, waited(seen), 'and the object must still be delivering'
  end

  # An Events that has had its last callback removed used to keep a parked
  # dispatcher thread and an entry the reader walks for every frame on the
  # connection -- measured: 50 proxies that registered one callback and
  # removed it left 51 live threads and 50 sink entries, until close. Nothing
  # is registered, so nothing is derived: that is the same rule the subscribe
  # and the Advise follow.
  def test_the_last_callback_removed_takes_the_dispatcher_and_the_sink_with_it
    client = FakeClient.new
    ev = WineOLE::Events.new(client, 1)
    assert_empty client.sinks, 'an ole_events nobody has registered on costs nothing'

    seen = Queue.new
    sub = ev.on('Click') { seen << :ran }
    assert_equal 1, client.sinks.length
    client.deliver(frame('Click', 1))
    assert_equal :ran, waited(seen), 'it must deliver before the off, or what follows proves nothing'

    ev.off(sub)
    assert ev.stopped_for_test?(5), 'the dispatcher must not stay parked with nothing to deliver'
    assert_empty client.sinks, 'and the sink must come off the connection'

    ev.on('Click') { seen << :again }
    client.deliver(frame('Click', 2))
    assert_equal :again, waited(seen), 'registering again must work exactly as the first time did'
  end

  # The bulk form, for a caller who does not want to remember the names.
  def test_close_takes_every_subscription_and_the_thread_down
    ev, client = events_for
    ev.on('Click') { }
    ev.on('Other') { }
    ev.close

    assert_equal 2, client.calls.count { |m, _| m == 'unsubscribe' }
    assert_empty ev.registered_names_for_test
    assert_empty client.sinks
    assert ev.stopped_for_test?(5)
  end

  def test_callbacks_run_in_registration_order
    ev, client = events_for
    order = Queue.new
    ev.on('Click') { order << :first }
    ev.on('Click') { order << :second }
    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 1, 'args' => nil})
    assert_equal :first, waited(order)
    assert_equal :second, waited(order)
  end

  def test_a_raising_callback_reaches_on_error_and_delivery_continues
    ev, client = events_for
    errors = Queue.new
    done = Queue.new
    ev.on_error { |e, _frame| errors << e }
    ev.on('Click') { raise 'boom' }
    ev.on('Click') { done << :ran }

    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 1, 'args' => nil})
    assert_equal 'boom', waited(errors)&.message
    assert_equal :ran, waited(done), 'a later callback must still run'

    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 2, 'args' => nil})
    assert_equal :ran, waited(done), 'and the next event must still be delivered'
  end

  # One connection can carry several objects with events. A registration
  # that REPLACED the previous consumer would switch the earlier object's
  # events off in silence.
  def test_two_event_objects_on_one_connection_both_receive
    client = FakeClient.new
    a = WineOLE::Events.new(client, 1)
    b = WineOLE::Events.new(client, 2)
    seen = Queue.new
    a.on('Click') { seen << :a }
    b.on('Click') { seen << :b }

    client.deliver({'event' => 'Click', 'handle' => 2, 'seq' => 1, 'args' => nil})
    assert_equal :b, waited(seen)
    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 2, 'args' => nil})
    assert_equal :a, waited(seen)
  end

  # The promise the README makes -- "one dispatcher thread per connection, in
  # arrival order, one at a time" -- and the whole value of it: a caller who
  # shares a Hash between an Application callback and a Workbook callback
  # needs no lock, because the two can never be inside their blocks at once.
  # Measured on a thread per Events object: both callbacks were in their
  # blocks simultaneously, on 2 distinct threads, with every other assertion
  # in this file still passing.
  def test_two_event_objects_on_one_connection_share_one_thread_and_never_overlap
    client = FakeClient.new
    a = WineOLE::Events.new(client, 1)
    b = WineOLE::Events.new(client, 2)
    entered = Queue.new
    gate = Queue.new
    threads = Queue.new
    a.on('Click') do
      threads << Thread.current
      entered << :a
      gate.pop
    end
    b.on('Click') do
      threads << Thread.current
      entered << :b
    end

    client.deliver(frame('Click', 1, handle: 1))
    assert_equal :a, waited(entered), 'the first callback must be inside its block'

    client.deliver(frame('Click', 2, handle: 2))
    assert_nil entered.pop(timeout: 0.5),
      'the second object\'s callback must not run while the first is held: one at a time'

    gate << :go
    assert_equal :b, waited(entered), 'and it must run as soon as the first returns'
    first_thread = waited(threads)
    second_thread = waited(threads)
    assert_same first_thread, second_thread,
      'both callbacks must have run on the connection\'s ONE dispatcher thread -- a thread ' \
      'per object is a data race in every callback that shares state with another object\'s'
  end

  # `off` means off, even for what is already in flight. A frame minted
  # before the unsubscribe reached the bridge names handles the bridge is
  # holding, so it is released rather than delivered -- dropping it silently
  # leaks two COM objects per event for the life of the connection.
  def test_a_frame_for_a_handle_with_no_target_is_released_not_delivered
    client = FakeClient.new
    ev = WineOLE::Events.new(client, 1)
    ran = Queue.new
    ev.on('Click') { ran << :ran }

    client.deliver(frame('Click', 42, args: [], handle: 2))
    ev.drain_for_test

    assert_empty ran, 'a frame names one object, and no other object may see it'
    assert_includes client.calls, ['release_event', {seq: 42}],
      'the handles the bridge minted for it must still go back'
  end

  # The connection's dispatcher belongs to the connection, not to whichever
  # object armed it first. One object going quiet must leave the others
  # exactly as they were -- same thread, same queue, no restart.
  def test_detaching_one_object_keeps_delivering_to_the_other_on_the_same_thread
    client = FakeClient.new
    a = WineOLE::Events.new(client, 1)
    b = WineOLE::Events.new(client, 2)
    seen = Queue.new
    threads = Queue.new
    a.on('Click') { seen << :a }
    b.on('Click') do
      threads << Thread.current
      seen << :b
    end

    client.deliver(frame('Click', 1, handle: 2))
    assert_equal :b, waited(seen), 'it must deliver before the off, or what follows proves nothing'
    before = waited(threads)

    a.off('Click')
    client.deliver(frame('Click', 2, handle: 2))
    assert_equal :b, waited(seen),
      "one object's last callback must not stop the connection's other deliveries"
    assert_same before, waited(threads), 'and it must still be the same dispatcher thread'
    refute b.stopped_for_test?(0.1), 'which is therefore still running'
  end

  # The idle hand-off, at its widest. A thread that has decided to go clears
  # its slot and then gives back what was left on the queue -- and giving one
  # frame back is a synchronous round trip, so the hand-off is (leftover
  # frames x RTT) wide, not nanoseconds. An object that attaches inside that
  # window gets a dispatcher of its own, and the frames it is sent go on the
  # SAME queue: the thread on its way out must not be able to pop one of
  # them. Measured on the code that drained outside the lock: two threads
  # alive at once, and a live subscription's event popped by the departing
  # one and given back instead of delivered -- a registered callback whose
  # event silently never arrives, which is the one state this class exists to
  # make unreachable.
  #
  # Every wait here is on evidence rather than on a clock: `wait_until_inside`
  # for the departing thread being inside its release, and `join` for it
  # having finished. That is what makes the interleaving the same one every
  # run.
  def test_a_frame_arriving_during_the_idle_hand_off_reaches_the_new_target
    client = OnceReleaseGatedClient.new
    a = WineOLE::Events.new(client, 1)
    in_callback = Queue.new
    hold = Queue.new
    a.on('Click') { in_callback << :in; hold.pop }
    departing = client.dispatcher.thread_for_test

    # Held inside a callback, which is the only way to build a queue up
    # behind the dispatcher.
    client.deliver(frame('Click', 1))
    assert_equal :in, waited(in_callback)

    # off, on, a frame, off. That sequence is the only way a frame ends up
    # BEHIND an :idle marker -- `detach` takes the sink off before pushing
    # one, so nothing can arrive behind it except through a later `attach`'s
    # sink. The frame in the middle is what the departing thread will have
    # left to give back, and giving it back is the round trip that holds it.
    a.off('Click')
    a.on('Click') { }
    client.deliver(frame('Click', 99, args: []))
    a.off('Click')

    hold << :go
    assert_equal 99, client.wait_until_inside,
      'the departing dispatcher must be inside the release of what it had left over'

    # It has cleared its thread slot by now, so this starts a second
    # dispatcher on the same connection and puts a new sink up.
    b = WineOLE::Events.new(client, 2)
    got = Queue.new
    entered_slow = Queue.new
    release_slow = Queue.new
    b.on('Slow') { entered_slow << :in; release_slow.pop }
    b.on('Click') { got << :b }
    client.deliver(frame('Slow', 78, handle: 2))
    assert_equal :in, waited(entered_slow), 'the new dispatcher must be running'

    # Queued while the new dispatcher is busy and the old one is still inside
    # its release: nobody is parked on the queue, so this frame sits there
    # for whichever thread reaches it first.
    client.deliver(frame('Click', 77, handle: 2))
    client.let_it_through
    assert departing.join(5), 'the departing dispatcher must finish'
    release_slow << :go

    assert_equal :b, waited(got),
      'a frame for an object that attached during the idle hand-off must reach it -- the ' \
      'thread on its way out must not pop it off the connection\'s queue'
    assert_equal [['release_event', {seq: 99}]],
      client.calls.select { |method, _| method == 'release_event' },
      'and what WAS left over from before the hand-off must still be given back'
  end

  # `@targets` is keyed by handle, and two `Events` on one handle would
  # otherwise unseat each other: the second `attach` overwrites the first's
  # slot, and either one's last `off` deletes it -- taking the other's
  # routing and the connection's dispatcher thread with it. Measured on the
  # code that overwrote: a frame for the shared handle reached only the
  # second object, and `first.off` then stopped the second's live
  # subscription and ended the dispatcher thread.
  #
  # `Proxy#ole_events` memoizes and bridge ids are unique per session, so
  # nothing shipped can reach this. But `Events.new(client, handle)` is
  # public, every test in this file builds them that way, and "removing the
  # wrong one silently stops a live consumer" is exactly what Client#off_event
  # refuses to leave to convention.
  def test_a_second_events_on_the_same_handle_is_refused_rather_than_unseating_the_first
    client = FakeClient.new
    first = WineOLE::Events.new(client, 1)
    second = WineOLE::Events.new(client, 1)
    seen = Queue.new
    first.on('Click') { seen << :first }

    error = assert_raises(ArgumentError) { second.on('Click') { seen << :second } }
    assert_match(/handle 1/, error.message, 'the refusal must name the handle it is about')

    assert_equal 1, client.sinks.length, 'the refused object must leave nothing on the connection'
    client.deliver(frame('Click', 2))
    assert_equal :first, waited(seen), 'the object that was there first must still receive'
    refute first.stopped_for_test?(0.1), "and the connection's dispatcher must still be running"
  end

  # A dispatcher blocked on an empty queue never exits, and every Events
  # leaks a thread for the life of the process.
  def test_the_dispatcher_thread_stops_when_the_connection_ends
    client = FakeClient.new
    ev = WineOLE::Events.new(client, 1)
    ev.on('Click') { }
    client.deliver(nil)
    assert ev.stopped_for_test?(2), 'the dispatcher must finish after the stream ends'
  end

  # The arguments are callback-scoped, so the release has to happen even when
  # the callback blew up -- otherwise one bad callback leaks two COM objects
  # per event, forever.
  def test_event_handles_are_released_even_when_the_callback_raises
    ev, client = events_for
    ev.on_error { }
    ev.on('Click') { raise 'boom' }
    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 77, 'args' => []})
    ev.drain_for_test
    assert_includes client.calls, ['release_event', {seq: 77}]
  end

  # A dead dispatcher is permanent and silent, which is why the rescue around
  # a callback is `Exception` and not the customary StandardError. Measured
  # on the code that used the narrow one: a callback raising outside
  # StandardError ended the thread, and the 10 events that followed were
  # queued and never delivered -- callbacks still registered, bridge still
  # advised, 1 of 11 release_events sent, on_error never told.
  def test_a_callback_raising_past_standard_error_does_not_stop_the_dispatcher
    ev, client = events_for
    errors = Queue.new
    ran = Queue.new
    ev.on_error { |e, _frame| errors << e }
    ev.on('Click') { raise Exception, 'past the usual rescue' }
    ev.on('Click') { ran << :later_callback }

    client.deliver(frame('Click', 78, args: []))
    assert_equal 'past the usual rescue', waited(errors)&.message, 'on_error must be told'
    assert_equal :later_callback, waited(ran), 'the callbacks after it must still run'

    client.deliver(frame('Click', 79, args: []))
    assert_equal :later_callback, waited(ran), 'and every later event must still be delivered'

    ev.drain_for_test
    assert_includes client.calls, ['release_event', {seq: 78}]
    assert_includes client.calls, ['release_event', {seq: 79}]
    refute ev.stopped_for_test?(0.1), 'the dispatcher thread must still be there'
  end

  # The other way to lose the thread, and the one the `ensure` in `deliver`
  # is really for: a frame whose `args` is neither an array nor null is
  # something the bridge cannot have sent, and it raises in build_args before
  # any callback is reached. The handles it names were minted all the same.
  def test_a_malformed_frame_reaches_on_error_and_still_releases_its_handles
    ev, client = events_for
    errors = Queue.new
    ran = Queue.new
    ev.on_error { |e, _frame| errors << e }
    ev.on('Click') { ran << :ran }

    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 80, 'args' => 'not an array'})
    assert_kind_of NoMethodError, waited(errors)
    ev.drain_for_test
    assert_includes client.calls, ['release_event', {seq: 80}],
      'the handles the bridge minted for it must still go back'

    client.deliver(frame('Click', 81, args: []))
    assert_equal :ran, waited(ran), 'and the next event must still be delivered'
  end

  # The third way it was lost: `report`'s own rescue was StandardError, so an
  # on_error raising past that killed the thread from inside the machinery
  # that exists to report exactly this.
  def test_an_on_error_that_itself_raises_does_not_stop_the_dispatcher
    ev, client = events_for
    ran = Queue.new
    ev.on_error { |_e, _frame| raise Exception, 'the handler is broken too' }
    ev.on('Click') { raise 'boom' }
    ev.on('Click') { ran << :later_callback }

    quiet do
      client.deliver(frame('Click', 1))
      assert_equal :later_callback, waited(ran), 'the callbacks after it must still run'
      client.deliver(frame('Click', 2))
      assert_equal :later_callback, waited(ran), 'and every later event must still be delivered'
    end
  end

  # Registering and keeping no reference is the ordinary shape --
  # `xl.ole_events.on('Click') { }` leaves the caller holding nothing but the
  # Proxy, and a caller who drops that holds nothing at all. The Client holds
  # the Dispatcher, and the Dispatcher's target table holds the Events: that
  # is the whole mechanism, and the opposite direction of the collectability
  # test below. Break it -- hold the targets weakly -- and the Events is
  # collected out from under a live connection: callbacks stop firing, the
  # bridge stays advised, and every event it goes on sending leaks its
  # argument handles because nobody is left to release them.
  def test_a_registered_callback_survives_a_garbage_collection
    client = FakeClient.new
    seen = Queue.new
    register_and_forget(client, seen)
    collect_garbage

    client.deliver({'event' => 'Click', 'handle' => 1, 'seq' => 1, 'args' => nil})
    assert_equal :ran, waited(seen), 'the callback must outlive the reference the caller dropped'
  end

  # In a method of its own so the Events is genuinely unreferenced -- a local
  # in the test body would keep it alive on the stack and the test would pass
  # whatever the sink closed over.
  def register_and_forget(client, seen)
    WineOLE::Events.new(client, 1).on('Click') { seen << :ran }
    nil
  end

  # A Client whose socket is never closed is the leak this phase already
  # measured once, at the reader thread, and fixed by making the Mailbox hold
  # its sinks weakly. A dispatcher thread started from a block written inside
  # an instance method walks it straight back in: the running thread captures
  # the block's binding, `self` included, so it roots the Dispatcher, which
  # holds the Client, whose finalizer therefore never runs. Measured on
  # exactly that code before this test existed -- the Client survived five
  # GC.starts and the peer never saw EOF.
  def test_a_client_with_events_on_it_is_still_collectable
    server, client_side = UNIXSocket.pair
    # Answers whatever `on` sends, so subscribing can complete. It holds the
    # server end only -- nothing that could pin the client under test.
    responder = Thread.new do
      begin
        while (line = server.gets)
          server.write(JSON.generate({id: JSON.parse(line)['id'], result: true}) + "\n")
        end
        :eof
      rescue IOError
        # The socket was closed under it by this test's own ensure, which
        # only happens on the failing path. Not a second failure to report.
        :closed
      end
    end

    dispatcher = nil
    # Client and Events are both built on a thread that is dead before the
    # question is asked, so this thread's stack never held a pointer to
    # either. `client = nil; ev = nil` here would not do it: assigning nil
    # does not unbuild the stack slots the collector has already seen. See
    # GarbageCollectionHelper.
    weak = weak_ref_to do
      client = WineOLE::Client.new(client_side)
      ev = WineOLE::Events.new(client, 1)
      ev.on('Click') { }
      # The Thread object, not `ev.method(:stopped_for_test?)`: a Method is
      # bound to its receiver, so carrying one out of this block would pin the
      # very Events this test is about. A Thread holds only what was handed to
      # it -- a WeakRef and the queue.
      dispatcher = client.dispatcher.thread_for_test
      client
    end

    assert_collected weak,
      'a Client with an Events on it must still be collectible -- something is pinning it'
    assert responder.join(5),
      'the peer must see EOF: the finalizer only closes the socket if the Client was collected'
    assert_equal :eof, responder.value
    # Deferred finalizers need the process to reach a safe point, so give it
    # ordinary work rather than parking on the thread straight away.
    100.times { |i| Array.new(100) { i.to_s } }
    assert dispatcher.join(5),
      'a collected Dispatcher must not leave its thread parked on a queue nothing can push to'
  ensure
    server&.close
  end

  # Closing the connection is the ordinary thing to do from a callback ("I
  # have seen what I was waiting for"), and it is a shutdown reaching in from
  # a thread the shutdown itself has to stop. It must return rather than
  # deadlock: on a real Client, over a real socket, with a real reader thread
  # to join.
  def test_a_callback_can_close_the_client_that_delivered_to_it
    server, client_side = UNIXSocket.pair
    responder = Thread.new do
      while (line = server.gets)
        request = JSON.parse(line)
        server.write(JSON.generate({id: request['id'], result: true}) + "\n")
        # The subscribe is answered, and then the event it subscribed to
        # arrives on the same stream, exactly as the bridge sends it.
        next unless request['method'] == 'subscribe'

        server.write(JSON.generate({event: 'Click', handle: 1, seq: 1, args: nil}) + "\n")
      end
    rescue IOError
      nil # the client closed the socket under it -- that is the test passing
    end

    client = WineOLE::Client.new(client_side)
    ev = WineOLE::Events.new(client, 1)
    closed = Queue.new
    ev.on('Click') do
      client.close
      closed << :returned
    end

    assert_equal :returned, waited(closed), 'close from inside a callback must return'
    assert responder.join(5)
  ensure
    server&.close
  end

  def quiet
    original = $stderr
    $stderr = StringIO.new
    yield
  ensure
    $stderr = original
  end
end
