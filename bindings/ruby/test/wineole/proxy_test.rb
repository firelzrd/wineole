require 'minitest/autorun'
require 'time'
require_relative '../../lib/wineole/proxy'

class FakeClient
  # A stand-in for the real Client#dispatcher, used only by the cleanup
  # tests below. It RECORDS a register_cleanup(handle, proc) call rather
  # than doing anything with it: the real Dispatcher#register_cleanup does
  # not exist yet (that lands in Task 8, which adds $cleanup DELIVERY), and
  # Proxy.register_cleanup must not depend on it to be unit-testable now.
  class RecordingDispatcher
    attr_reader :registered_cleanups

    def initialize
      @registered_cleanups = []
    end

    def register_cleanup(handle, proc)
      @registered_cleanups << [handle, proc]
    end
  end

  attr_reader :calls, :await_cleanup_calls
  attr_accessor :connect_or_create_created, :release_reply

  def initialize
    @calls = []
    @connect_or_create_created = true
    @event_sinks = []
    @release_reply = nil
    @await_cleanup_calls = []
  end

  # Lazily, like the real Client -- a test that never touches `cleanup:`
  # should not have to care that this exists.
  def dispatcher
    @dispatcher ||= RecordingDispatcher.new
  end

  def await_cleanup(seq)
    @await_cleanup_calls << seq
  end

  # Proxy#ole_events builds an Events, and an Events registers a sink here
  # and starts a dispatcher thread. `end_stream` is what lets a test that
  # touched ole_events leave no thread parked behind it.
  def on_event(&block)
    @event_sinks << block
    self
  end

  def end_stream
    @event_sinks.each { |sink| sink.call(nil) }
  end

  def call(method, params)
    @calls << [method, params]
    case method
    when 'create' then {'$ole_ref' => 1}
    when 'connect' then {'$ole_ref' => 1}
    when 'connect_or_create' then {'$ole_ref' => 1, 'created' => @connect_or_create_created}
    when 'invoke'
      case params[:name]
      when 'Version' then 11.0
      when 'Worksheets' then {'$ole_ref' => 2}
      when 'LastSaveTime' then {'$type' => 'time', 'iso8601' => '2024-03-05T12:30:00'}
      when 'BulkValue'
        [
          [1, {'$type' => 'time', 'iso8601' => '2026-08-31T09:30:00'}],
          [nil, 'text'],
        ]
      when 'BulkRefs' then [{'$ole_ref' => 42}, {'$ole_ref' => 43}]
      when 'NestedHash' then {'when' => {'$type' => 'time', 'iso8601' => '2026-08-31T09:30:00'}}
      else nil
      end
    when 'release' then @release_reply
    when 'leave_open' then nil
    when 'const_load' then {'xlUp' => -4162, 'xlDown' => -4121}
    end
  end
end

class ProxyTest < Minitest::Test
  def test_method_missing_forwards_invoke_and_decodes_primitive
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    assert_equal 11.0, proxy.Version

    method, params = client.calls.last
    assert_equal 'invoke', method
    assert_equal 1, params[:handle]
    assert_equal 'Version', params[:name]
    assert_equal [], params[:args]
    assert_equal({}, params[:named])
  end

  def test_method_missing_decodes_tagged_time_into_ruby_time
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    result = proxy.LastSaveTime
    assert_instance_of Time, result
    assert_equal Time.iso8601('2024-03-05T12:30:00'), result
  end

  def test_method_missing_decodes_ole_ref_into_new_proxy
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    sheets = proxy.Worksheets
    assert_instance_of WineOLE::Proxy, sheets
    assert_equal 2, sheets.ole_handle
  end

  def test_trailing_hash_argument_is_sent_as_named_args
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    sheets = proxy.Worksheets

    sheets.Add('After' => sheets)

    _method, params = client.calls.last
    assert_equal 'Add', params[:name]
    assert_equal [], params[:args]
    assert_equal({'After' => {'$ole_ref' => 2}}, params[:named])
  end

  def test_default_indexer_uses_empty_name
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    sheets = proxy.Worksheets

    sheets[1]

    _method, params = client.calls.last
    assert_equal '', params[:name]
    assert_equal [1], params[:args]
  end

  def test_marshal_dump_raises_not_serializable
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    assert_raises(WineOLE::NotSerializableError) { Marshal.dump(proxy) }
  end

  def test_ole_events_is_the_same_object_every_time
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    # Memoized, not rebuilt: an Events owns a dispatcher thread and the set
    # of subscriptions the bridge knows about, so an `off` reaching a
    # different object than the `on` did would leave the subscription up and
    # the callback unreachable.
    assert_same proxy.ole_events, proxy.ole_events
    client.end_stream
  end

  def test_ole_events_raises_on_a_stale_reference
    client = FakeClient.new
    other_client = FakeClient.new
    stale = WineOLE::Proxy.wrap(client, other_client.object_id, 1)

    # Same rule as every other member here: a handle from a previous
    # connection means nothing on this one, and subscribing to events on it
    # would advise some unrelated object -- or nothing at all.
    assert_raises(WineOLE::StaleReferenceError) { stale.ole_events }
    assert_empty client.calls
  end

  def test_stale_reference_raises_without_calling_client
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    other_client = FakeClient.new
    stale = WineOLE::Proxy.wrap(client, other_client.object_id, 1)

    assert_raises(WineOLE::StaleReferenceError) { stale.Version }
    assert_equal 1, client.calls.length # only the initial 'create' call — no 'invoke' attempted
  end

  # --- Ruby's implicit conversion protocol must not become RPC traffic ---

  def test_implicit_conversion_hooks_are_not_forwarded_to_com
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    baseline = client.calls.length

    %i[to_ary to_a to_hash to_str to_io to_proc to_int to_i to_f to_path coerce].each do |hook|
      assert_raises(NoMethodError, "#{hook} should behave like a plain object's missing method") do
        proxy.public_send(hook)
      end
      refute proxy.respond_to?(hook), "respond_to?(#{hook}) must be false"
    end

    assert_equal baseline, client.calls.length,
      'no implicit-conversion hook may cost a remote call'
  end

  def test_integer_and_numeric_coercion_do_not_call_the_client
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    baseline = client.calls.length

    assert_raises(TypeError) { Integer(proxy) }
    assert_raises(TypeError) { 1 + proxy }

    assert_equal baseline, client.calls.length,
      'Integer() and numeric coercion must not produce remote calls'
  end

  def test_respond_to_still_true_for_ordinary_com_members
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    assert proxy.respond_to?(:Version)
    assert proxy.respond_to?(:Worksheets)
  end

  def test_puts_and_p_and_array_and_destructuring_do_not_call_the_client
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    baseline = client.calls.length

    # `puts`/`p` probe to_ary; Array() probes to_ary then to_a; multiple
    # assignment probes to_ary. All used to be forwarded to the remote
    # object and blow up with DISP_E_UNKNOWNNAME after a wasted round trip.
    devnull = File.open(File::NULL, 'w')
    begin
      devnull.puts(proxy)          # IO#puts probes to_ary, then to_s
      devnull.print(proxy.inspect) # `p` probes inspect
      assert_equal [proxy], Array(proxy)
      a, b = proxy
      assert_same proxy, a
      assert_nil b
      assert_kind_of String, "#{proxy}"
    ensure
      devnull.close
    end

    assert_equal baseline, client.calls.length,
      'printing/converting a Proxy must not produce remote calls'
  end

  def test_proxy_from_a_different_client_cannot_be_passed_as_an_argument
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    sheets = proxy.Worksheets

    other_client = FakeClient.new
    foreign = WineOLE::Proxy.create('Excel.Application', other_client)

    baseline = client.calls.length
    assert_raises(WineOLE::StaleReferenceError) { sheets.Add('After' => foreign) }
    assert_equal baseline, client.calls.length,
      'a foreign handle must never be encoded onto this connection'

    # ...including nested inside arrays/hashes.
    assert_raises(WineOLE::StaleReferenceError) { sheets.Add([foreign]) }
    assert_equal baseline, client.calls.length
  end

  def test_ole_const_load_returns_the_raw_hash_from_the_wire
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    consts = proxy.ole_const_load

    assert_equal({'xlUp' => -4162, 'xlDown' => -4121}, consts)
    method, params = client.calls.last
    assert_equal 'const_load', method
    assert_equal 1, params[:handle]
  end

  def test_ole_const_load_raises_on_a_stale_reference_without_calling_the_client
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    other_client = FakeClient.new
    stale = WineOLE::Proxy.wrap(client, other_client.object_id, 1)

    assert_raises(WineOLE::StaleReferenceError) { stale.ole_const_load }
    assert_equal 1, client.calls.length # only the initial 'create' call
  end

  def test_proxy_create_reports_created_true
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    assert_equal true, proxy.ole_created?
  end

  def test_proxy_connect_reports_created_false
    client = FakeClient.new
    proxy = WineOLE::Proxy.connect('Excel.Application', client)
    assert_equal false, proxy.ole_created?
  end

  def test_proxy_connect_or_create_reports_created_from_the_wire
    client = FakeClient.new
    client.connect_or_create_created = false
    proxy = WineOLE::Proxy.connect_or_create('Excel.Application', client)

    assert_equal false, proxy.ole_created?
    method, params = client.calls.last
    assert_equal 'connect_or_create', method
    assert_equal 'Excel.Application', params[:class_name]
  end

  def test_proxy_wrap_reports_created_as_nil
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    sheets = proxy.Worksheets

    assert_nil sheets.ole_created?
  end

  def test_invoke_raises_on_a_stale_reference_without_calling_the_client
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)
    other_client = FakeClient.new
    stale = WineOLE::Proxy.wrap(client, other_client.object_id, 1)

    assert_raises(WineOLE::StaleReferenceError) { stale.invoke('Version', [], {}) }
    assert_equal 1, client.calls.length # only the initial 'create' call
  end

  def test_invoke_is_a_public_escape_hatch
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    result = proxy.invoke('Version', [], {})

    assert_equal 11.0, result
    method, params = client.calls.last
    assert_equal 'invoke', method
    assert_equal 'Version', params[:name]
  end

  def test_decode_converts_values_nested_inside_an_array
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    rows = proxy.BulkValue

    assert_equal 1, rows[0][0]
    assert_instance_of Time, rows[0][1],
      'a date inside a bulk range read must decode to a Time, not stay a raw Hash'
    assert_equal Time.iso8601('2026-08-31T09:30:00'), rows[0][1]
    assert_nil rows[1][0]
    assert_equal 'text', rows[1][1]
  end

  def test_decode_converts_ole_refs_nested_inside_an_array
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    sheets = proxy.BulkRefs

    assert_instance_of WineOLE::Proxy, sheets[0]
    assert_equal 42, sheets[0].ole_handle
    assert_equal 43, sheets[1].ole_handle
  end

  def test_decode_converts_values_nested_inside_a_plain_hash
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    result = proxy.NestedHash

    assert_instance_of Time, result['when']
  end

  def test_encode_converts_a_time_to_the_wire_tag
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    proxy.invoke('Value=', [Time.new(2026, 8, 31, 9, 30, 0)], {})

    _method, params = client.calls.last
    assert_equal(
      [{'$type' => 'time', 'iso8601' => '2026-08-31T09:30:00'}],
      params[:args],
      'a Time must go out as the same tag the receive side emits'
    )
  end

  def test_encode_converts_times_nested_inside_an_array
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    proxy.invoke('Value=', [[[Time.new(2026, 8, 31, 9, 30, 0), 1]]], {})

    _method, params = client.calls.last
    assert_equal(
      [[[{'$type' => 'time', 'iso8601' => '2026-08-31T09:30:00'}, 1]]],
      params[:args],
      'encode is recursive, so a date inside a bulk write must be tagged too'
    )
  end

  def test_encode_converts_a_date_to_the_wire_tag_at_midnight
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    proxy.invoke('Value=', [Date.new(2026, 8, 31)], {})

    _method, params = client.calls.last
    assert_equal(
      [{'$type' => 'time', 'iso8601' => '2026-08-31T00:00:00'}],
      params[:args],
      'a bare Date must go out as midnight'
    )
  end

  def test_encode_converts_a_datetime_to_the_wire_tag_preserving_its_time
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    proxy.invoke('Value=', [DateTime.new(2026, 8, 31, 9, 30, 0)], {})

    _method, params = client.calls.last
    assert_equal(
      [{'$type' => 'time', 'iso8601' => '2026-08-31T09:30:00'}],
      params[:args],
      'a DateTime must keep its time, proving the Date branch also covers DateTime'
    )
  end

  # --- cleanup:, ole_leave_open, and ole_release's cleanup-key reply ---

  def test_create_sends_converted_cleanup_steps_and_callback_flag
    client = FakeClient.new
    on_cleanup = proc {}

    # on_cleanup present => callback true; steps [name,*args] => {name:,args:}
    proxy = WineOLE::Proxy.create('Excel.Application', client,
      cleanup: {steps: [['DisplayAlerts=', false], ['Quit']], on_cleanup: on_cleanup})

    method, params = client.calls.first
    assert_equal 'create', method
    assert_equal 'Excel.Application', params[:class_name]
    assert_equal(
      {steps: [{name: 'DisplayAlerts=', args: [false]}, {name: 'Quit', args: []}], callback: true},
      params[:cleanup]
    )
    assert_equal [[proxy.ole_handle, on_cleanup]], client.dispatcher.registered_cleanups,
      'the client closure must be registered against the created handle'
  end

  def test_create_without_on_cleanup_sends_callback_false_and_registers_nothing
    client = FakeClient.new

    WineOLE::Proxy.create('Excel.Application', client, cleanup: {steps: [['Quit']]})

    _method, params = client.calls.first
    assert_equal({steps: [{name: 'Quit', args: []}], callback: false}, params[:cleanup])
    assert_empty client.dispatcher.registered_cleanups
  end

  def test_create_without_a_cleanup_argument_sends_no_cleanup_key
    client = FakeClient.new

    WineOLE::Proxy.create('Excel.Application', client)

    _method, params = client.calls.first
    refute params.key?(:cleanup), 'a caller that never asked for cleanup must not get a cleanup key on the wire'
  end

  def test_ole_leave_open_sends_the_wire_call_and_returns_nil
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    assert_nil proxy.ole_leave_open

    method, params = client.calls.last
    assert_equal 'leave_open', method
    assert_equal 1, params[:handle]
  end

  def test_ole_release_returns_nil
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    assert_nil proxy.ole_release
  end

  def test_ole_release_awaits_cleanup_when_the_bridge_replies_with_a_cleanup_key
    client = FakeClient.new
    client.release_reply = {'cleanup' => 5}
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    assert_nil proxy.ole_release
    assert_equal [5], client.await_cleanup_calls
  end

  def test_ole_release_does_not_await_when_the_bridge_replies_with_no_cleanup_key
    client = FakeClient.new
    proxy = WineOLE::Proxy.create('Excel.Application', client)

    proxy.ole_release
    assert_empty client.await_cleanup_calls
  end
end
