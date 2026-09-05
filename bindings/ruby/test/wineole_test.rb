require 'minitest/autorun'
require_relative '../lib/wineole'

class WineOLETest < Minitest::Test
  def teardown
    WineOLE.close
  end

  def test_default_client_lazy_init_is_thread_safe_and_calls_open_only_once
    original_open = WineOLE::Client.method(:open)
    open_call_count = 0
    count_mutex = Mutex.new
    fake_client = Object.new
    def fake_client.create(class_name) = class_name
    def fake_client.close = nil

    WineOLE::Client.define_singleton_method(:open) do |**_kwargs|
      count_mutex.synchronize { open_call_count += 1 }
      sleep 0.05 # widen the race window so concurrent callers actually overlap
      fake_client
    end

    results = []
    results_mutex = Mutex.new
    threads = Array.new(20) do
      Thread.new do
        result = WineOLE.create('Excel.Application')
        results_mutex.synchronize { results << result }
      end
    end
    threads.each(&:join)

    assert_equal 1, open_call_count, 'Client.open must be called exactly once no matter how many threads race'
    assert_equal ['Excel.Application'], results.uniq
  ensure
    WineOLE::Client.define_singleton_method(:open, original_open)
  end

  def test_open_updates_the_implicit_default
    fake_client_a = Object.new
    def fake_client_a.create(class_name) = "a:#{class_name}"
    def fake_client_a.close = nil

    original_open = WineOLE::Client.method(:open)
    WineOLE::Client.define_singleton_method(:open) { |**_kwargs| fake_client_a }

    opened = WineOLE.open
    assert_same fake_client_a, opened

    result = WineOLE.create('Excel.Application')
    assert_equal 'a:Excel.Application', result, 'WineOLE.create must use the client WineOLE.open just set as the default'
  ensure
    WineOLE::Client.define_singleton_method(:open, original_open)
  end

  def test_close_clears_the_default_so_the_next_create_opens_a_fresh_one
    fake_client_a = Object.new
    def fake_client_a.create(class_name) = class_name
    def fake_client_a.close = nil
    fake_client_b = Object.new
    def fake_client_b.create(class_name) = class_name
    def fake_client_b.close = nil

    original_open = WineOLE::Client.method(:open)
    clients = [fake_client_a, fake_client_b]
    WineOLE::Client.define_singleton_method(:open) { |**_kwargs| clients.shift }

    WineOLE.create('Excel.Application')
    first_default = WineOLE.send(:default_client)
    assert_same fake_client_a, first_default

    WineOLE.close
    WineOLE.create('Excel.Application')
    second_default = WineOLE.send(:default_client)

    assert_same fake_client_b, second_default
    refute_same first_default, second_default, 'close must force the next .create to open a fresh client'
  ensure
    WineOLE::Client.define_singleton_method(:open, original_open)
  end

  def test_default_client_is_public
    assert WineOLE.public_methods.include?(:default_client),
      'default_client must be public so bundled wrappers like MSOffice::Excel can reach ' \
      'the one connection this module already owns, instead of opening a second one'
  end

  def test_default_client_returns_the_lazily_initialized_client
    fake_client = Object.new
    def fake_client.close = nil

    original_open = WineOLE::Client.method(:open)
    WineOLE::Client.define_singleton_method(:open) { |**_kwargs| fake_client }

    result = WineOLE.default_client
    assert_same fake_client, result
  ensure
    WineOLE::Client.define_singleton_method(:open, original_open)
  end

  def test_connect_or_create_uses_the_default_client
    fake_client = Object.new
    def fake_client.connect_or_create(class_name) = "coc:#{class_name}"
    def fake_client.close = nil

    original_open = WineOLE::Client.method(:open)
    WineOLE::Client.define_singleton_method(:open) { |**_kwargs| fake_client }

    result = WineOLE.connect_or_create('Excel.Application')
    assert_equal 'coc:Excel.Application', result
  ensure
    WineOLE::Client.define_singleton_method(:open, original_open)
  end
end
