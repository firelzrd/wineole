require_relative 'errors'

module WineOLE
  # COM events for one object, reached as `proxy.ole_events`.
  #
  # Registering a callback is the only thing a caller does. Everything under
  # it is DERIVED from that. The bridge-side subscription and the COM Advise
  # beneath it belong to THIS object: they are put up by its first callback
  # and taken back down by the last one removed from it. The sink on the
  # connection and the dispatcher thread belong to the CONNECTION, so they
  # follow the same rule one level up -- up with the first callback anywhere
  # on the client, down with the last one anywhere on it (see Dispatcher).
  # Making any of them a separate control would allow the state where a
  # callback is registered and the event never arrives, with nothing to show
  # for it.
  #
  # That invariant is a claim about ORDER as much as about bookkeeping, so
  # `on`, `off` and `close` hold @wire_mutex across "decide, then tell the
  # bridge". Deciding under one lock and writing the wire outside it lets two
  # threads reach the bridge in the opposite order to the one they decided
  # in: a subscribe landing after the unsubscribe that was meant to follow it
  # leaves a registered callback whose event can never arrive, and the mirror
  # case leaves the bridge advised with nothing to deliver to. Measured, on
  # the code before this lock existed: wire order subscribe, subscribe,
  # unsubscribe with one callback still registered.
  #
  # Callbacks run on one dispatcher thread PER CONNECTION, in arrival order,
  # one at a time -- so two objects on one client never run their callbacks
  # concurrently, and a caller needs no lock between them. That thread is not
  # here: it belongs to the connection, and this object attaches itself to it
  # for as long as it has a registration (see Dispatcher). They may call COM
  # freely -- that is the point -- because the client's reader thread is a
  # different thread and stays free to read the response. A slow callback
  # delays the events behind it, on every object on the connection.
  #
  # The dispatcher thread survives anything a callback or a malformed frame
  # can throw; see Dispatcher.run.
  class Events
    # NOT a Struct, and not for tidiness: a Struct has value equality, so two
    # subscriptions with the same name, block and flag are `==` without being
    # the same registration -- registering one proc twice for an event is
    # enough. `off(first)` would then remove the second one too (Array#delete
    # removes every match), unsubscribe, and leave the caller with a
    # Subscription that never fires again and nothing said about it. A
    # registration is the thing itself, so identity is its equality.
    #
    # `args` is per callback, but the wire has one flag per event: what goes
    # on it is the union (see `effective_args`).
    class Subscription
      attr_reader :name, :block, :args

      def initialize(name, block, args)
        @name = name
        @block = block
        @args = args
      end
    end

    def initialize(client, handle)
      @client = client
      @handle = handle
      # A plain Hash, deliberately not `Hash.new { |h, k| h[k] = [] }`: with a
      # default block, merely LOOKING an event name up writes a permanent key
      # -- an arriving event nobody subscribed to, or an `off` for a name that
      # was never `on`, would grow this table for the life of the connection.
      # A name is added when a callback is registered for it and removed when
      # the last one goes.
      @subs = {}
      @mutex = Mutex.new
      # Held across the whole of `on`/`off`/`close` -- the decision AND the
      # call that carries it out -- and always acquired BEFORE @mutex, never
      # while holding it. @mutex is therefore never held across a wire round
      # trip, which is what keeps `deliver`'s brief acquisition of it off the
      # bridge's response time. A callback calling `off` from the dispatcher
      # waits for at most one round trip and cannot deadlock: the thread it
      # waits for needs nothing from the dispatcher to finish.
      #
      # The connection's Dispatcher has a mutex of its own, and it is taken
      # under this one and never under @mutex: the order is @wire_mutex ->
      # Dispatcher#@mutex (in `arm`/`disarm`) and @wire_mutex -> @mutex
      # (everywhere else), so those two are never held together. The
      # Dispatcher takes none of these, and drops its own before it calls
      # `deliver`, which is what lets a callback register on another object.
      @wire_mutex = Mutex.new
      @on_error = nil
      # Whether this object currently has a place on the connection's
      # dispatcher. It stands where @sink used to: the dispatcher and its
      # sink are the connection's now, shared with every other object on it,
      # so "is it up?" is no longer a question this object can answer by
      # looking at what it owns. Read and written under @wire_mutex only --
      # `arm` and `disarm` are both called with it held -- which is what
      # keeps `disarm` idempotent, and it has to be: it is reached from
      # `off`'s ensure, from `close`, and from `on`'s rescue, on an object
      # that may never have been armed at all.
      @attached = false
    end

    # `args: false` tells the bridge not to mint handles for this event's
    # object arguments. The callback is then called with no arguments at all
    # -- measured: a block written `|sheet, range|` gets nil for both. Worth
    # it for a high-frequency event you only want to count.
    #
    # The bridge holds ONE flag per event, so what goes on the wire is the
    # union: arguments are minted while any callback for that event wants
    # them. Registering an `args: true` callback next to an `args: false` one
    # therefore re-subscribes rather than leaving the first registration's
    # flag standing -- measured on Excel before this was here: the second
    # callback was handed nil, having asked for the objects, and nothing said
    # so. A callback that asked for `args: false` and gets them anyway
    # because a sibling wanted them is the harmless direction of the same
    # trade; it can ignore them.
    def on(name, args: true, &block)
      raise ArgumentError, 'on needs a block -- there is nothing to call otherwise' unless block

      sub = Subscription.new(name, block, args)
      @wire_mutex.synchronize do
        # Before the subscribe, never after: the bridge advises the COM
        # source as the subscribe is handled, so an event can be on its way
        # back before the call returns. Arming afterwards would drop it.
        arm
        wanted = @mutex.synchronize do
          before = effective_args(name)
          (@subs[name] ||= []) << sub
          after = effective_args(name)
          after == before ? nil : after
        end
        begin
          @client.call('subscribe', {handle: @handle, event: name, args: wanted}) unless wanted.nil?
        rescue StandardError
          # A subscribe the bridge refused -- an object that is not an event
          # source is the ordinary case -- must not leave the callback
          # registered. Keeping it would produce exactly the state this class
          # exists to make unreachable: a callback that is never called, with
          # the error already raised and gone.
          disarm if @mutex.synchronize { drop(name, sub); @subs.empty? }
          raise
        end
      end
      sub
    end

    # Takes either a name (every callback for it) or one Subscription.
    def off(name_or_sub)
      @wire_mutex.synchronize do
        sub = name_or_sub.is_a?(Subscription) ? name_or_sub : nil
        name = sub ? sub.name : name_or_sub.to_s
        before, after, empty = @mutex.synchronize do
          was = effective_args(name)
          drop(name, sub)
          [was, effective_args(name), @subs.empty?]
        end
        # Nothing was registered for this event, so nothing was derived from
        # it either. `after.nil?` alone cannot tell "the last callback just
        # went" from "there was never one", and an unsubscribe for a
        # subscription that does not exist is a round trip that says nothing.
        next if before.nil?

        begin
          if after.nil?
            # The last callback for this event is gone, so the subscription
            # that only existed to feed it goes too -- and with the last name
            # for the object, the COM Advise underneath it.
            @client.call('unsubscribe', {handle: @handle, event: name})
          elsif after != before
            # Callbacks remain, but the one that wanted arguments was among
            # those removed: stop paying for handles nobody asked for.
            @client.call('subscribe', {handle: @handle, event: name, args: after})
          end
        ensure
          # In an ensure because the registry has already been emptied above:
          # if the unsubscribe raised (a connection that has just gone is the
          # ordinary case) the local half must still come down, or this
          # object keeps a dispatcher thread and a sink entry on the client
          # for the life of the connection with nothing left to deliver.
          disarm if empty
        end
      end
      self
    end

    # Every callback forgotten, every subscription and Advise released, the
    # dispatcher stopped and the sink taken off the connection.
    #
    # `off`-ing the last callback does all of this already -- that is the
    # derivation this class is built on, and it is why there is no `close`
    # you are REQUIRED to call. This is the bulk form, for a caller who does
    # not want to remember which names it registered. `on` afterwards works
    # exactly as the first one did: the object arms itself again.
    def close
      @wire_mutex.synchronize do
        @mutex.synchronize { @subs.keys }.each do |name|
          @mutex.synchronize { drop(name, nil) }
          begin
            @client.call('unsubscribe', {handle: @handle, event: name})
          rescue StandardError
            # A connection that has already gone has already unadvised
            # everything on it. Unlike `off`, this is not reported: `close`
            # is what a caller reaches for when it is done, and the local
            # half it exists to release comes down either way.
            nil
          end
        end
        disarm
      end
      self
    end

    # ONE error handler per object, last writer wins -- deliberately not
    # `on`'s append. An error handler is not a subscription: it has no
    # arguments to negotiate, nothing is derived from it, and there is
    # nothing for a second one to add that the first cannot do. That it
    # returns `self` rather than a Subscription is the same statement, and
    # there is no `off_error` for the same reason -- `on_error { }` replaces
    # it with one that does nothing.
    def on_error(&block)
      @mutex.synchronize { @on_error = block }
      self
    end

    # Tests only. Production code never needs any of these, because callbacks
    # are the delivery mechanism.
    #
    # The queue and the thread these two ask about belong to the connection
    # now, so both are the Dispatcher's answers. Kept here as delegators
    # because a test that has an `Events` and asks whether ITS callbacks are
    # still being delivered is asking the right question of the right object
    # -- and because the answer is the same one it always was.
    def stopped_for_test?(seconds)
      @client.dispatcher.stopped_for_test?(seconds)
    end

    def drain_for_test(seconds = 5)
      @client.dispatcher.drain_for_test(seconds)
      self
    end

    def registered_names_for_test
      @mutex.synchronize { @subs.keys }
    end

    private

    # Takes this object's place on the connection: the frames for its handle
    # start being routed to it, and the connection's dispatcher thread and
    # sink go up if this is the first object to ask for them. Derived from
    # there being a callback at all, so an `ole_events` a caller merely
    # touched costs nothing anywhere. Called from `on` under @wire_mutex,
    # which is what serializes it against `disarm`.
    def arm
      return if @attached

      @client.dispatcher.attach(@handle, self)
      @attached = true
    end

    # The other half of `arm`, once the last callback is gone: a registered
    # callback is the only reason for any of this to exist. Without it an
    # Events that has had every callback removed still holds a place on the
    # connection's dispatcher and an entry the reader walks for every frame
    # on it -- measured: 50 proxies that registered one callback and removed
    # it left 51 live threads and 50 sink entries, all of them until the
    # connection closed. The thread and the sink themselves only go with the
    # LAST object to leave; that decision belongs to the Dispatcher, which is
    # the only thing that knows whether any other object still has callbacks.
    def disarm
      return unless @attached

      @attached = false
      # By identity: the Dispatcher removes this object's routing entry and
      # nobody else's.
      @client.dispatcher.detach(@handle, self)
    end

    # What the wire flag for `name` should be: nil when no callback is
    # registered for it at all, otherwise true if any of them wants the
    # event's object arguments minted. Call it under @mutex.
    def effective_args(name)
      subs = @subs[name]
      return nil if subs.nil? || subs.empty?

      subs.any? { |sub| sub.args }
    end

    # Removes one subscription, or every callback for `name` when `sub` is
    # nil, and takes the name out of the table when nothing is left for it.
    # Call it under @mutex.
    #
    # `equal?`, not `==`: `off` removes the registration it was handed and no
    # other. See Subscription.
    def drop(name, sub)
      list = @subs[name]
      return if list.nil?

      sub.nil? ? list.clear : list.reject! { |s| s.equal?(sub) }
      @subs.delete(name) if list.empty?
    end

    # One frame, to the callbacks registered for its name. Called by the
    # connection's Dispatcher, which routed it here by handle, and private
    # for the reason the whole class exists: a registered callback is the
    # only way in, and a public `deliver` would say otherwise. The frame's
    # argument handles are given back by the Dispatcher's `route`, in an
    # ensure that covers a frame this raised on as well as one that reached
    # nobody at all.
    def deliver(frame)
      subs = @mutex.synchronize { (@subs[frame['event']] || []).dup }
      args = build_args(frame['args'])
      subs.each do |sub|
        begin
          sub.block.call(*args)
        rescue Exception => e # rubocop:disable Lint/RescueException
          # Everything, for the reason Dispatcher.run catches everything: this
          # thread is the whole delivery mechanism, and a callback raising
          # something outside StandardError must not take the next callback,
          # the next event and every later release down with it.
          report(e, frame)
        end
      end
    end

    def build_args(raw)
      return [] if raw.nil?

      # The same decode an invoke's result goes through, so an event argument
      # that is an object arrives as a Proxy and one that is a date arrives
      # as a Time -- a second, private copy of that walk here is how those
      # two answers drift apart.
      raw.map { |v| Proxy.decode(@client, v) }
    end

    # Never raises. It is called from the dispatcher's own rescue, so an
    # exception out of here is the one thing that could still end the thread.
    def report(err, frame)
      handler = @mutex.synchronize { @on_error }
      event = frame.is_a?(Hash) ? frame['event'] : nil
      begin
        if handler
          handler.call(err, frame)
        else
          warn "wineole: #{event} callback raised #{err.class}: #{err.message}"
        end
      rescue Exception => e # rubocop:disable Lint/RescueException
        # An on_error that itself raises must not recurse, and must not be
        # able to kill the dispatcher either -- its own rescue was
        # StandardError once, and an on_error raising past that was a third
        # way to lose every later event in silence.
        warn "wineole: on_error raised #{e.class} while reporting #{err.class}"
      end
    rescue Exception # rubocop:disable Lint/RescueException
      nil # even $stderr being gone must not end the dispatcher
    end
  end
end
