require 'time'
require 'date'
require_relative 'errors'
# Stated here rather than left to wineole.rb: `ole_events` below names the
# constant, so anything that loads this file alone -- proxy_test.rb does --
# must get a working method rather than a NameError the first time it is
# called. events.rb requires nothing from here (it reaches Proxy at call
# time, not load time), so there is no cycle.
require_relative 'events'

module WineOLE
  class Proxy
    def self.create(class_name, client, cleanup: nil)
      params = {class_name: class_name}
      params[:cleanup] = build_cleanup(cleanup) if cleanup
      handle = client.call('create', params)['$ole_ref']
      register_cleanup(client, handle, cleanup)
      new(client, session_id: client.object_id, handle: handle, created: true)
    end

    def self.connect(class_name, client, cleanup: nil)
      params = {class_name: class_name}
      params[:cleanup] = build_cleanup(cleanup) if cleanup
      handle = client.call('connect', params)['$ole_ref']
      register_cleanup(client, handle, cleanup)
      new(client, session_id: client.object_id, handle: handle, created: false)
    end

    def self.connect_or_create(class_name, client, cleanup: nil)
      params = {class_name: class_name}
      params[:cleanup] = build_cleanup(cleanup) if cleanup
      result = client.call('connect_or_create', params)
      handle = result['$ole_ref']
      register_cleanup(client, handle, cleanup)
      new(client, session_id: client.object_id, handle: handle, created: result['created'])
    end

    # `cleanup` is `{ steps: [[name, *args], ...], on_cleanup: proc_or_nil }`
    # in Ruby terms; the wire wants `{ steps: [{name:, args:}, ...], callback: bool }`.
    # `callback` tells the bridge whether to hold the root open and emit a
    # `$cleanup` event for a registered closure, or just run the steps and
    # release outright -- so its value comes from whether `on_cleanup` is
    # present, not from anything the caller states separately.
    def self.build_cleanup(cleanup)
      steps = (cleanup[:steps] || []).map do |name, *args|
        {name: name, args: args}
      end
      {steps: steps, callback: !cleanup[:on_cleanup].nil?}
    end
    private_class_method :build_cleanup

    # Register the client closure (if any) so the dispatcher can deliver
    # $cleanup for this root handle. The real Dispatcher#register_cleanup
    # lands in a later task; this only needs `client.dispatcher` to answer to
    # it, which is exactly what the real Client already does.
    def self.register_cleanup(client, handle, cleanup)
      return unless cleanup && cleanup[:on_cleanup]

      client.dispatcher.register_cleanup(handle, cleanup[:on_cleanup])
    end
    private_class_method :register_cleanup

    # One wire value, in Ruby terms.
    #
    # On the class rather than private to an instance because two different
    # holders of a client need it: an invoke's RESULT, and an event's
    # ARGUMENTS (WineOLE::Events#build_args). A second copy of this walk over
    # there is how the same tagged value ends up a Time when a call returns it
    # and a raw {"$type" => "time"} Hash when an event carries it -- and how a
    # nested $ole_ref reaches a callback unwrapped.
    #
    # Recursive, because a bulk range read comes back as an array of rows and
    # the values needing conversion sit inside it, not at the top level. A
    # non-recursive decode would hand back raw {"$type" => "time"} hashes for
    # every date cell in the range.
    def self.decode(client, value)
      case value
      when Array
        value.map { |v| decode(client, v) }
      when Hash
        if value.key?('$ole_ref')
          wrap(client, client.object_id, value['$ole_ref'])
        elsif value['$type'] == 'time'
          Time.iso8601(value['iso8601'])
        else
          value.transform_values { |v| decode(client, v) }
        end
      else
        value
      end
    end

    def self.wrap(client, session_id, ole_ref)
      new(client, session_id: session_id, handle: ole_ref, created: nil)
    end

    # Ruby's implicit-conversion protocol. The interpreter probes these on
    # arbitrary objects behind the scenes — `puts`/`p` and `Array()` look for
    # `to_ary`/`to_a`, multiple assignment looks for `to_ary`, string
    # interpolation and `Kernel#String` look for `to_str`, `IO` methods look
    # for `to_io`, `&obj` looks for `to_proc`, `Integer()`/`format('%d', _)`/
    # `"ab" * _`/array indexing look for `to_int`/`to_i`, `Float()` looks for
    # `to_f`, `File.open` looks for `to_path`, and `1 + _` looks for `coerce`.
    # Forwarding these to the remote object turns every `puts proxy` into a
    # round trip that ends in DISP_E_UNKNOWNNAME, so they must behave exactly
    # as they would on a plain Ruby object that doesn't define them:
    # NoMethodError, and `respond_to?` == false (which is what stops the
    # interpreter calling them at all).
    IMPLICIT_CONVERSIONS = %i[
      to_ary to_a to_hash to_str to_io to_proc
      to_int to_i to_f to_path coerce
    ].freeze

    attr_reader :ole_handle, :ole_session_id

    def initialize(client, session_id:, handle:, created:)
      @client = client
      @ole_session_id = session_id
      @ole_handle = handle
      @created = created
    end

    def method_missing(name, *args)
      return super if IMPLICIT_CONVERSIONS.include?(name)

      check_live!
      named = args.last.is_a?(Hash) ? args.pop : {}
      invoke(name.to_s, args, named)
    end

    def respond_to_missing?(name, _include_private = false)
      !IMPLICIT_CONVERSIONS.include?(name)
    end

    def [](*args)
      check_live!
      invoke('', args, {})
    end

    # Was this instance freshly created by connect_or_create's
    # fallback, or attached to something already running? `true` for
    # `.create`, `false` for `.connect`, whatever the bridge reported for
    # `.connect_or_create`, and `nil` for anything derived from another
    # Proxy (e.g. `xl.Worksheets`) — attach-vs-create isn't a meaningful
    # question for those.
    def ole_created?
      @created
    end

    # Tells the bridge to leave this root instance running rather than
    # closing it when the connection goes away -- the counterpart to a
    # `cleanup:` closure a caller wants to run later, on its own terms,
    # rather than as part of this process's teardown.
    def ole_leave_open
      @client.call('leave_open', {handle: @ole_handle})
      nil
    end

    def ole_release
      result = @client.call('release', {handle: @ole_handle})
      # A client closure must run before this handle is actually gone: the
      # bridge answers with the $cleanup sequence number instead of
      # releasing outright, and this blocks until the dispatcher has
      # delivered it and the release_event that follows. See
      # Client#await_cleanup.
      if result.is_a?(Hash) && (seq = result['cleanup'])
        @client.await_cleanup(seq)
      end
      nil
    end

    # COM events for this object. Named with the `ole_` prefix like every
    # other bookkeeping method here: a Proxy forwards unknown names straight
    # to COM, so a bare `events` would shadow a real `Events` member.
    #
    # Memoized, because the Events owns a dispatcher thread and a bridge-side
    # subscription set: a fresh one per call would mean `on` and the `off`
    # that is meant to undo it talked to different objects.
    def ole_events
      check_live!
      @ole_events ||= Events.new(@client, @ole_handle)
    end

    def ole_const_load
      check_live!
      @client.call('const_load', {handle: @ole_handle})
    end

    def marshal_dump
      raise NotSerializableError, 'WineOLE::Proxy references are connection-scoped and cannot be persisted'
    end

    # Deliberately bare and public, unlike every other meta-method here
    # (which are `ole_`-prefixed to avoid shadowing a same-named remote COM
    # member): an explicit escape hatch for the rare case a COM object
    # really does define e.g. an `ole_handle` member, matching real Ruby
    # WIN32OLE's own choice to keep `invoke` public and unprefixed.
    def invoke(name, args, named)
      check_live!
      params = {
        handle: @ole_handle,
        name: name,
        args: args.map { |a| encode(a) },
        named: named.transform_values { |v| encode(v) },
      }
      decode(@client.call('invoke', params))
    end

    private

    def check_live!
      return if @ole_session_id == @client.object_id

      raise StaleReferenceError, 'this reference belongs to a previous connection'
    end

    def encode(value)
      case value
      when Proxy
        # The argument's own liveness is not enough: a Proxy belonging to a
        # *different* Client is live from its own point of view, yet its
        # handle id means nothing in this connection's handle table (or,
        # worse, means something unrelated). Check it against the receiver's
        # client, which is the connection the id is about to be sent over.
        unless value.ole_session_id == @client.object_id
          raise StaleReferenceError,
            'this reference belongs to a different connection and cannot be ' \
            'passed as an argument here'
        end
        {'$ole_ref' => value.ole_handle}
      when Time
        # The same tag the receive side emits for VT_DATE. The wall clock is
        # sent as-is: a VT_DATE carries no timezone, so converting here would
        # silently move the value the caller wrote. Matching
        # wineole/proxy.py's `_encode`.
        {'$type' => 'time', 'iso8601' => value.strftime('%Y-%m-%dT%H:%M:%S')}
      when Date
        # `date` is already loaded transitively -- `time`'s own lib/time.rb
        # requires it -- so this require costs nothing new; it just states
        # the real dependency instead of relying on another library's
        # internals. DateTime is a subclass of Date, and both answer
        # strftime, so one branch covers them. (Date is not in Time's
        # hierarchy, which is why the `when Time` branch above doesn't catch
        # it.) Matching wineole/proxy.py's `_encode`.
        {'$type' => 'time', 'iso8601' => value.strftime('%Y-%m-%dT%H:%M:%S')}
      when Hash
        value.transform_values { |v| encode(v) }
      when Array
        value.map { |v| encode(v) }
      else
        value
      end
    end

    def decode(value)
      Proxy.decode(@client, value)
    end
  end
end
