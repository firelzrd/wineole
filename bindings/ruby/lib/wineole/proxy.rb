require 'time'
require_relative 'errors'

module WineOLE
  class Proxy
    def self.create(class_name, client)
      handle = client.call('create', {class_name: class_name})['$ole_ref']
      new(client, session_id: client.object_id, handle: handle, created: true)
    end

    def self.connect(class_name, client)
      handle = client.call('connect', {class_name: class_name})['$ole_ref']
      new(client, session_id: client.object_id, handle: handle, created: false)
    end

    def self.connect_or_create(class_name, client)
      result = client.call('connect_or_create', {class_name: class_name})
      new(client, session_id: client.object_id, handle: result['$ole_ref'], created: result['created'])
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

    def ole_release
      @client.call('release', {handle: @ole_handle})
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
      when Hash
        value.transform_values { |v| encode(v) }
      when Array
        value.map { |v| encode(v) }
      else
        value
      end
    end

    def decode(value)
      if value.is_a?(Hash) && value.key?('$ole_ref')
        Proxy.wrap(@client, @client.object_id, value['$ole_ref'])
      elsif value.is_a?(Hash) && value['$type'] == 'time'
        Time.iso8601(value['iso8601'])
      else
        value
      end
    end
  end
end
