require 'weakref'
require_relative 'errors'

module WineOLE
  # The one dispatcher thread of a CONNECTION, and the queue it parks on.
  #
  # One thread per connection is a promise the README makes to the caller
  # ("Callbacks run on one dispatcher thread per connection, in arrival
  # order, one at a time"), and the whole value of it is that a caller never
  # needs a lock BETWEEN callbacks: a Hash shared by an Application callback
  # and a Workbook callback is safe because the two can never be inside their
  # blocks at the same time. A thread per `Events` object breaks exactly that
  # and nothing else -- measured, with two `Events` on one connection: both
  # callbacks were inside their blocks simultaneously, on 2 distinct threads,
  # while every other event assertion in the suite still passed.
  #
  # Everything here is DERIVED from there being a registered callback
  # somewhere on the connection, the same way `Events` derives its
  # subscription and its Advise from one: the thread and the sink go up at
  # the first `attach` and come back down after the last `detach`. An
  # `ole_events` nobody registered on costs no thread, and a connection whose
  # callbacks have all been removed goes back to costing none.
  #
  # THE MUTEX IS NEVER HELD WHILE USER CODE RUNS. A callback that registers a
  # callback on ANOTHER object -- `sheet.ole_events.on(...)` from inside an
  # Application callback -- reaches `attach` on this very thread, and would
  # deadlock against a lock its own delivery still held.
  class Dispatcher
    def initialize(client)
      @client = client
      # handle -> Events, at most ONE Events per handle per connection. That
      # invariant is what makes a handle a sufficient routing key, and
      # `Proxy#ole_events` is what guarantees it in shipped code: it memoizes,
      # and a bridge id is unique per session, so two Proxies never share a
      # handle. `Events.new(client, handle)` is public, though, so `attach`
      # enforces it rather than trusting it -- a second Events silently
      # unseating the first would leave a registered callback that never
      # fires, and either object's last `off` would then stop the other.
      #
      # Held STRONGLY and on purpose. This is what keeps an
      # `Events` alive exactly as long as it has registrations, even when the
      # caller kept no reference to the Proxy it came from
      # (`xl.ole_events.on('Click') { }` leaves the caller holding nothing).
      # The chain is Client -> Dispatcher -> @targets -> Events. Hold them
      # weakly and the Events is collected out from under a live connection:
      # callbacks stop firing, the bridge stays advised, and every event it
      # goes on sending leaks its argument handles because nobody is left to
      # release them. Nothing here pins the Client from outside, because the
      # Client holds this Dispatcher and the Events holds the Client -- the
      # whole ring is collected together or not at all.
      @targets = {}
      # handle -> on_cleanup closure. Registered by a Proxy on the on_cleanup
      # path; a $cleanup frame for that handle runs the closure here, on the
      # dispatcher thread, the same COM-safe context every other callback runs
      # on. Kept beside @targets because a client that uses on_cleanup but
      # subscribes to no events still needs the sink and thread up, so this
      # counts toward "is there a registered callback on the connection" the
      # same way @targets does.
      @cleanups = {}
      @mutex = Mutex.new
      @queue = Queue.new
      @sink = nil
      @thread = nil
      # A collected Dispatcher must not leave its thread parked on a queue
      # nothing can ever push to again. The finalizer pushes the same :stop
      # the end of the stream does, and captures the QUEUE ONLY: capturing
      # `self` would keep this object reachable from ObjectSpace's finalizer
      # table forever, so it could never be collected and the finalizer could
      # never run -- the same rule as Client.finalizer.
      ObjectSpace.define_finalizer(self, self.class.stopper(@queue))
    end

    def self.stopper(queue)
      proc { queue << :stop }
    end

    # The connection's one sink, built on the class so it captures the queue
    # and nothing else. It is registered once, at the first `attach`, and
    # does no filtering: every frame on the connection goes on the one queue
    # and is routed by handle on the dispatcher thread, which is what makes
    # "in arrival order, one at a time" true across objects rather than
    # within one.
    #
    # This runs on the reader thread -- or, for a client whose stream has
    # already ended, on the thread that is registering right now. It must
    # enqueue and return either way: it may not block and it may not raise.
    def self.sink(queue)
      proc do |frame|
        # nil means the connection is gone.
        queue << (frame.nil? ? :stop : frame)
      end
    end

    # Registers `events` as the target for its handle, and puts up whatever
    # the connection does not have yet. Called from `Events#arm` under that
    # object's @wire_mutex; two `Events` on one connection have two of those,
    # so @mutex is what makes this safe between them.
    #
    # The calls out of this class made under @mutex are `Client#on_event`
    # here and `Client#off_event` in `detach`, and both are safe by
    # inspection: neither runs user code (the sink above is ours, and only
    # enqueues), and the locks they take -- the Client's sink list and the
    # Mailbox's -- are ordered against @mutex, not disjoint from it. This
    # class does wait for the Mailbox's lock, in `release`, but only with
    # @mutex NOT held; and nothing holding either of those locks ever waits
    # for @mutex: `Client#on_event` lets its sink list go before it calls
    # the Mailbox, and the Mailbox drops its lock before it calls a sink.
    # The one thing either can do on this thread -- `on_event` handing an
    # end-of-stream nil straight back when the stream has already ended --
    # only pushes :stop on the queue.
    #
    # That hand-off reaches the FIRST object to arm on a dead connection and
    # not a second one, because by then the sink is already registered: the
    # second object's thread parks on a queue nothing will push to until the
    # Dispatcher's finalizer pushes :stop as the ring is collected. Bounded,
    # and no ordinary path reaches it -- `client.close` still ends the
    # stream through the one sink that is up.
    def attach(handle, events)
      @mutex.synchronize do
        current = @targets[handle]
        if current && !current.equal?(events)
          # Only a DIFFERENT object is refused: the caller holding two
          # `Events` for one object. The same one is let through so that a
          # re-attach can never be mistaken for that (an ordinary `on` after
          # `close` finds the slot empty, since `detach` cleared it). See
          # @targets.
          raise ArgumentError,
                "handle #{handle} already has an Events on this connection; one object's " \
                'events belong to one Events (Proxy#ole_events memoizes for this reason)'
        end

        @targets[handle] = events
        if @sink.nil?
          @sink = Dispatcher.sink(@queue)
          @client.on_event(&@sink)
        end
        next if @thread&.alive?

        # `&Dispatcher.method(:run)`, not `{ run_loop }`. A block captures its
        # whole binding, `self` included, so a thread started from a block
        # written in an instance method is a permanent GC root for this
        # Dispatcher -- and through @client, for the Client. Measured on
        # exactly that code, one layer down: the Client was never collected,
        # its finalizer never ran, and its socket stayed open for the life of
        # the process. That is the same leak the Mailbox was extracted to fix
        # (see Client::Mailbox), walked back in through this thread. A Method
        # object on the class captures the class and nothing else, so what
        # the thread holds is a WeakRef and a queue.
        @thread = Thread.new(WeakRef.new(self), @queue, &Dispatcher.method(:run))
        @thread.abort_on_exception = false
      end
      self
    end

    # The other half of `attach`, for one object. The connection's thread and
    # sink only come down with the LAST one: an Application whose callbacks
    # are all removed must not stop the Workbook's events on the same
    # connection.
    def detach(handle, events)
      @mutex.synchronize do
        # By identity, and for the reason `Client#off_event` uses identity:
        # removing a routing entry that belongs to somebody else silently
        # stops a live consumer. With `attach` refusing a second Events on
        # one handle this can only be a stale detach -- `disarm` on an object
        # that has already left -- but "it cannot happen" is not a reason to
        # write the unguarded delete.
        @targets.delete(handle) if @targets[handle].equal?(events)
        sink = @sink
        # A registered cleanup keeps the sink and thread up even with no
        # events left: the $cleanup frame it is waiting for still has to be
        # delivered on this connection's dispatcher.
        next unless @targets.empty? && @cleanups.empty? && sink

        @sink = nil
        @client.off_event(sink)
        # After the sink comes off, never before: the marker means "nothing
        # more arrives behind this through the sink", and a frame the reader
        # pushed between the two is behind the marker, drained by
        # `confirm_idle` and released. Not quite airtight: a reader ALREADY
        # inside `dispatch_to_sinks` when `off_event` ran still holds the old
        # sink, and can push after the drain too. Such a frame stays on the
        # queue with no thread, until the next `attach` starts one -- which
        # delivers it if that object is back for the event, and releases it
        # otherwise. Bounded by that one dispatch, and unseen in a 400-cycle
        # detach/attach stress (2572 frames), but it is there. Never a join,
        # either -- the thread reaching here is very often the dispatcher
        # itself, in a callback that called `off`.
        @queue << :idle
      end
      self
    end

    # Register a client closure to run when the bridge asks (a $cleanup frame
    # for `handle`). Arms the connection's sink and thread the same way
    # `attach` does, so a client that uses on_cleanup but subscribes to no
    # events still has a dispatcher to deliver the frame on. Runs no user code
    # under @mutex -- storing a closure and installing our own sink are the
    # only things done here -- so holding it across the whole body is safe,
    # exactly as it is in `attach`.
    def register_cleanup(handle, block)
      @mutex.synchronize do
        @cleanups[handle] = block
        if @sink.nil?
          @sink = Dispatcher.sink(@queue)
          @client.on_event(&@sink)
        end
        next if @thread&.alive?

        # A Method object on the class, never a block: see `attach` for why a
        # block started as a thread here is a permanent GC root for the Client.
        @thread = Thread.new(WeakRef.new(self), @queue, &Dispatcher.method(:run))
        @thread.abort_on_exception = false
      end
      self
    end

    # The other half of `register_cleanup`. Brings the sink and thread down
    # with the `@queue << :idle` hand-off, exactly as `detach`'s last-target
    # case does -- but only when nothing is left to deliver to: a cleanup
    # removed while events are still registered must not stop them, and vice
    # versa, which is why the guard is the same `@targets.empty? &&
    # @cleanups.empty?` teardown condition.
    def unregister_cleanup(handle)
      @mutex.synchronize do
        @cleanups.delete(handle)
        sink = @sink
        next unless @targets.empty? && @cleanups.empty? && sink

        @sink = nil
        @client.off_event(sink)
        @queue << :idle
      end
      self
    end

    # The dispatcher thread's body, on the class so it captures no instance.
    # The Dispatcher is reached weakly and never held across a `pop`, which
    # is what lets a Client with events on it still be collected.
    #
    # THIS THREAD SURVIVES EVERYTHING. A dead dispatcher is permanent and
    # silent: the callbacks stay registered, the bridge stays advised, every
    # later event's argument handles leak on the bridge for the life of the
    # connection, the queue grows without bound, and `on_error` is never
    # told. Measured on the code before this rescue existed -- a callback
    # raising outside StandardError, or a frame whose `args` was not an
    # array -- 10 later events were queued and never delivered, and 1 of 11
    # release_events was sent. So the rescue is `Exception`, not
    # StandardError: the narrow rescue exists to let a fatal exception reach
    # a thread that can act on it, and there is no such thread here -- this
    # one dying takes the whole feature down without a word. Reported through
    # `report`, which cannot raise, and the loop goes on to the next frame.
    def self.run(ref, queue)
      while (item = queue.pop)
        break if item == :stop

        dispatcher = begin
          ref.__getobj__
        rescue WeakRef::RefError
          break # there is nobody left to deliver to
        end

        begin
          break if step(dispatcher, item)
        rescue Exception => e # rubocop:disable Lint/RescueException
          dispatcher.__send__(:report, e, item)
        end
        # Dropped before parking again: a local still pointing at the
        # Dispatcher would pin it for as long as this thread waits, which is
        # precisely what the WeakRef is here to avoid.
        dispatcher = nil
      end
    ensure
      # However this thread ended, the Dispatcher must not go on believing it
      # has one: `attach` decides whether to start a thread from exactly that.
      begin
        ref.__getobj__.__send__(:thread_finished, Thread.current)
      rescue WeakRef::RefError
        nil
      end
    end

    # One queued item. Answers whether the thread is to end. On the class,
    # like `run`, so it captures no instance -- and it no longer needs the
    # queue: what is left on it when this thread goes is taken off under
    # @mutex by `confirm_idle`, which is the whole of the fix that made the
    # hand-off safe.
    def self.step(dispatcher, item)
      case item
      when :idle
        # `detach` left this behind when the last target went away. It only
        # ends the thread if that is still true: an `attach` that got in
        # first answers no and the SAME thread carries on, which is what
        # keeps "one dispatcher, arrival order" true across a detach/attach
        # cycle instead of briefly running two.
        #
        # When the answer is yes, `confirm_idle` hands back everything that
        # was still queued -- taken off the queue under the same lock, at the
        # same instant, as the thread slot was cleared. Releasing it is a
        # round trip per frame, so it is done HERE, out of that lock: holding
        # @mutex across a round trip would stall every attach and detach on
        # the connection behind it. A thread that has decided to go can
        # therefore still be RELEASING while its successor runs, but it can
        # no longer take anything off the queue, which is what it must not
        # do: a frame pushed after that instant came from a later `attach`'s
        # sink, and that sink was installed under the same lock.
        leftover = dispatcher.__send__(:confirm_idle)
        return false if leftover.nil?

        # Whatever happens in here, the answer is "end": the slot is already
        # clear, so a successor may be running, and a raise escaping to `run`'s
        # rescue would keep THIS thread on the queue beside it. `release`
        # rescues StandardError itself; this is for the rest.
        begin
          leftover.each { |left| dispatcher.__send__(:release, left) }
        rescue Exception => e # rubocop:disable Lint/RescueException
          dispatcher.__send__(:report, e, item)
        end
        true
      when Array
        item.last.call if item.first == :barrier
        false
      when Hash
        # Real frames are always Hashes here (:idle/:stop/[:barrier,..] are the
        # non-Hash items). A $cleanup frame goes to the client's on_cleanup
        # closure and is acked; every other frame is routed by handle to its
        # Events, exactly as before.
        if item['event'] == '$cleanup'
          dispatcher.__send__(:run_cleanup, item)
        else
          dispatcher.__send__(:route, item)
        end
        false
      end
    end

    # Is `thread` this dispatcher's own callback thread? Used by
    # Client#await_cleanup to avoid a self-wait when `ole_release` is called
    # from inside a callback -- the $cleanup frame that release triggers is
    # queued behind the very callback asking the question, so waiting for it
    # here would deadlock the dispatcher against itself.
    def on_thread?(thread)
      @mutex.synchronize { @thread }&.equal?(thread) || false
    end

    # Tests only. Production code never needs any of these, because callbacks
    # are the delivery mechanism.
    def stopped_for_test?(seconds)
      thread = @mutex.synchronize { @thread }
      return true if thread.nil?

      !thread.join(seconds).nil?
    end

    # The Thread itself, for the collectability tests: a Thread holds only
    # what was handed to it -- a WeakRef and the queue -- so carrying one out
    # of a `weak_ref_to` block pins neither this Dispatcher nor its Client.
    def thread_for_test
      @mutex.synchronize { @thread }
    end

    # Blocks until the dispatcher has finished everything queued so far.
    # Bounded, because the dispatcher not finishing is exactly what a test
    # using this is hunting: an unbounded wait would hang the suite instead
    # of reporting it.
    #
    # A Mutex/ConditionVariable barrier rather than the shorter
    # `Queue#pop(timeout:)`: that keyword arrived in Ruby 3.2 and this gem
    # declares `required_ruby_version >= 3.0` (wineole.gemspec), so on the
    # oldest Ruby it supports the shorter form is an ArgumentError raised
    # from shipped code.
    def drain_for_test(seconds = 5)
      done = false
      lock = Mutex.new
      cond = ConditionVariable.new
      @queue << [:barrier, -> { lock.synchronize { done = true; cond.broadcast } }]

      deadline = Process.clock_gettime(Process::CLOCK_MONOTONIC) + seconds
      lock.synchronize do
        until done
          left = deadline - Process.clock_gettime(Process::CLOCK_MONOTONIC)
          raise Error, "the dispatcher did not drain within #{seconds}s" if left <= 0

          cond.wait(lock, left)
        end
      end
      self
    end

    private

    # A $cleanup frame: run the client closure for its handle on THIS thread
    # (COM-safe, like every other callback), then tell the bridge the closure
    # is done and wake whoever is blocked in await_cleanup. The closure's own
    # exception must not stop any of that -- the bridge runs the steps
    # regardless (choice B), so a raising closure still ends with release_event
    # and the waiter signalled.
    #
    # The closure is looked up UNDER @mutex and called WITHOUT it: like a
    # target's callback, it is free to `on`/`off`/`leave_open` anything on this
    # connection, each of which re-enters this class and would deadlock against
    # a lock its own delivery still held.
    def run_cleanup(frame)
      handle = frame['handle']
      seq = frame['seq']
      block = @mutex.synchronize { @cleanups[handle] }
      begin
        block&.call
      rescue StandardError => e
        warn "wineole: on_cleanup closure raised #{e.class}: #{e.message}"
      ensure
        begin
          @client.call('release_event', {seq: seq})
        rescue StandardError
          nil
        end
        @client.signal_cleanup_done(seq)
        @mutex.synchronize { @cleanups.delete(handle) }
      end
    end

    # One frame, to the object it names. The target is looked up under
    # @mutex and called WITHOUT it: `deliver` runs user code, and the
    # callback is free to `on` or `off` anything on this connection, which
    # comes straight back here as an `attach` or a `detach`.
    #
    # `__send__` rather than making `Events#deliver` public: delivering is
    # not something a caller may do -- that a registered callback is the only
    # way in is the whole claim of that class -- and a public `deliver` would
    # say otherwise in the one place a user reads. The Dispatcher is the
    # other half of the same feature, and this is the only place it reaches
    # across.
    def route(frame)
      target = target_for(frame)
      # A frame for a handle with no target was minted before the
      # unsubscribe reached the bridge. Not delivered -- `off` means off --
      # but still released by the ensure below, because the COM objects
      # behind those handles would otherwise sit on the bridge until the
      # connection closed.
      target.__send__(:deliver, frame) if target
    ensure
      # The arguments are valid for the callback and no longer. Releasing in
      # an ensure is what makes that true even when the frame never reached a
      # callback at all -- a malformed `args` raises inside `deliver` before
      # any callback is reached, and its handles would otherwise leak for the
      # life of the connection.
      release(frame)
    end

    # The frames of one event are released together, whether they reached a
    # callback, reached one that raised, reached no target at all, or were
    # still queued when the last target left. One statement of the rule --
    # `route`'s ensure and the `:idle` arm both come here -- because "when is
    # there something to give back" answered in two places is how the two
    # answers drift apart.
    def release(frame)
      # `args: null` says the bridge minted NOTHING for this event: it is
      # serialized as null rather than left out precisely so the client can
      # tell that from "this event had zero arguments" (protocol.rs), and
      # nothing is inserted in the bridge's event table for it. Sending a
      # release anyway is a synchronous round trip from the dispatcher, which
      # caps the event rate at one RTT -- in exactly the `args: false` case a
      # caller reaches for to keep up with a high-frequency event.
      return if !frame.is_a?(Hash) || frame['args'].nil?

      @client.call('release_event', {seq: frame['seq']})
    rescue StandardError
      # The connection is going away; there is nothing left to release.
      nil
    end

    def target_for(item)
      return nil unless item.is_a?(Hash)

      @mutex.synchronize { @targets[item['handle']] }
    end

    # Never raises. It is called from the dispatcher's own rescue, so an
    # exception out of here is the one thing that could still end the thread.
    # Routed to the object the frame names, so that an `on_error` registered
    # on it is told; a frame that names nobody has no handler to reach.
    def report(err, item)
      target = target_for(item)
      if target
        target.__send__(:report, err, item)
      else
        warn "wineole: dispatcher raised #{err.class}: #{err.message}"
      end
    rescue Exception # rubocop:disable Lint/RescueException
      nil # even $stderr being gone must not end the dispatcher
    end

    # The dispatcher's question when it reaches an :idle marker: is it still
    # true that there is nothing to deliver to? Answering it under the same
    # lock `attach` decides in is what makes "start a thread only if there is
    # none" safe -- either this clears @thread first and `attach` starts a
    # fresh one, or `attach` registers its target first and this thread keeps
    # running.
    #
    # nil when the answer is no; otherwise everything still on the queue,
    # drained HERE rather than by the caller. That is the whole point of it
    # being here: clearing the thread slot lets an `attach` start a
    # successor, and the successor's sink is installed under this same lock,
    # so a queue emptied in the same critical section as the slot is cleared
    # provably holds nothing of the successor's. Draining afterwards instead,
    # outside the lock, is what let a departing thread pop a frame belonging
    # to an object that had just attached and give it back rather than
    # deliver it -- measured: two dispatcher threads alive at once, and a
    # live subscription's event released, never delivered.
    def confirm_idle
      @mutex.synchronize do
        # A registered cleanup counts the same as a target: the thread must
        # stay to deliver the $cleanup frame it is waiting for.
        next nil unless @targets.empty? && @cleanups.empty?

        @thread = nil
        drain_queue
      end
    end

    # Everything on the queue, without blocking on it -- which is what makes
    # it safe to call under @mutex. Whatever comes back was minted before the
    # unsubscribe reached the bridge, with no callback left to hand it to.
    def drain_queue
      leftover = []
      loop { leftover << @queue.pop(true) }
    rescue ThreadError
      leftover # the queue is empty, which is the only way out of that loop
    end

    def thread_finished(thread)
      @mutex.synchronize { @thread = nil if @thread.equal?(thread) }
    end
  end
end
