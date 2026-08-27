require 'minitest/autorun'
require_relative '../../lib/wineole/proxy'

class FakeClient
  attr_reader :calls
  attr_accessor :connect_or_create_created

  def initialize
    @calls = []
    @connect_or_create_created = true
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
      else nil
      end
    when 'release' then nil
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
end
