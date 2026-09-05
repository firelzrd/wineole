require 'socket'
require 'json'
require 'rbconfig'
require 'tmpdir'
require 'ipaddr'
require 'weakref'
require_relative 'errors'
require_relative 'proxy'
require_relative 'dispatcher'

module WineOLE
  class Client
    WINDOWS = !(RbConfig::CONFIG['host_os'] =~ /mswin|mingw|cygwin/).nil?

    ARCH_TRIPLES = {
      'x86_64' => 'x86_64-pc-windows-gnu',
      'i386' => 'i686-pc-windows-gnu',
      'i686' => 'i686-pc-windows-gnu',
    }.freeze

    def self.bridge_path_for_arch(host_cpu)
      triple = ARCH_TRIPLES[host_cpu] || (ARCH_TRIPLES['i686'] if host_cpu =~ /^i.86$/)
      if triple.nil?
        raise Error, "no prebuilt wineole-bridge binary for host architecture #{host_cpu.inspect} " \
                      "(available: #{ARCH_TRIPLES.values.uniq.join(', ')})"
      end
      File.expand_path("../../wineole-bridge-dist/#{triple}/wineole-bridge.exe", __dir__)
    end

    def self.default_bridge_path
      bridge_path_for_arch(RbConfig::CONFIG['host_cpu'])
    end

    DEFAULT_SPAWNER = lambda do |port|
      # wineole-bridge.exe is a native Windows binary either way -- `wine`
      # is only needed to run it on a non-Windows host. On Windows itself
      # there is no `wine` command, so running it under one would fail
      # immediately (Errno::ENOENT) rather than run natively as it should.
      # The null device path is also platform-specific ('/dev/null' does
      # not exist on Windows).
      command = WINDOWS ? [default_bridge_path] : ['wine', default_bridge_path]
      null_device = WINDOWS ? 'NUL' : '/dev/null'
      Process.spawn(*command, port.to_s, %i[out err] => null_device)
    end

    def self.default_lockfile(port)
      # The lock is per-port, not global: two bridges on two ports are
      # entirely independent, and a single shared lock would make a client
      # starting one wait on a client starting the other (design doc §7.1
      # step 2). Dir.tmpdir rather than a hardcoded '/tmp' resolves
      # correctly on Windows (%TEMP%) as well as every POSIX platform.
      File.join(Dir.tmpdir, "wineole-bridge.#{port}.lock")
    end

    def self.open(host: '127.0.0.1', port: 47800, spawner: DEFAULT_SPAWNER,
                   lockfile: nil, timeout: 15, token: nil)
      lockfile ||= default_lockfile(port)

      socket = try_connect(host, port)
      return handshake(new(socket), token) if socket

      File.open(lockfile, File::CREAT | File::RDWR, 0o644) do |lock|
        lock.flock(File::LOCK_EX)

        socket = try_connect(host, port)
        return handshake(new(socket), token) if socket

        spawner.call(port)
        deadline = Time.now + timeout
        loop do
          socket = try_connect(host, port)
          return handshake(new(socket), token) if socket
          raise Error, "wineole-bridge did not start within #{timeout}s" if Time.now > deadline
          sleep 0.2
        end
      end
    end

    # Design doc §7.1 step 1: ping after connecting, to confirm protocol
    # compatibility before using the connection. This is also the only place a
    # token can be presented — a bridge started with WINEOLE_TOKEN rejects
    # every non-loopback request until a matching ping has been accepted, so
    # without this a tokened bridge is unreachable from this client.
    #
    # A failed ping (bad token, incompatible bridge) raises out of
    # `Client#call` as a RemoteError; it must not be swallowed.
    def self.handshake(client, token)
      client.call('ping', token ? {token: token} : {})
      client
    rescue StandardError
      client.close
      raise
    end
    private_class_method :handshake

    def self.try_connect(host, port)
      socket = TCPSocket.new(host, port)
      # Mirrors the bridge's own set_nodelay (main.rs). Insurance only: the
      # ~40 ms per-RPC stall this project hit was on the response side, and
      # requests already go out in a single write, so this changes nothing
      # today -- it keeps a future multi-write request path from
      # reintroducing it.
      socket.setsockopt(Socket::IPPROTO_TCP, Socket::TCP_NODELAY, 1)
      socket
    rescue Errno::ECONNREFUSED, Errno::ETIMEDOUT, Errno::EHOSTUNREACH
      nil
    end
    private_class_method :try_connect

    # A proc to close over the raw socket for ObjectSpace.define_finalizer.
    # It deliberately does NOT close over `self` (the Client instance) --
    # capturing the instance being finalized would keep it permanently
    # reachable through ObjectSpace's own finalizer table, and it could
    # never actually be collected. Capturing only the socket avoids this.
    def self.finalizer(socket)
      proc do
        begin
          socket.close
        rescue StandardError
          nil
        end
      end
    end

    def initialize(socket)
      @socket = socket
      # The one strong reference to the event sinks; see `on_event`. Written
      # there and never read -- the reader thread reaches the sinks through
      # the Mailbox's weak references instead, so this list is what keeps them
      # alive. Delete it as dead code and a plain `client.on_event { }` keeps
      # working only until the next garbage collection, after which the
      # connection silently stops delivering events. The guard is
      # ClientEventsTest#test_a_registered_consumer_survives_a_garbage_collection.
      @event_sinks = []
      @event_sinks_mutex = Mutex.new
      # Eagerly, and deliberately: a Dispatcher is a Hash, a Queue and two
      # nil slots until something attaches to it -- no thread, no sink, no
      # socket traffic -- so building it here costs less than the lock that
      # building it lazily would need, on a connection that may have several
      # objects registering callbacks at once.
      @dispatcher = Dispatcher.new(self)
      @mailbox = Mailbox.new(socket).start
      # Backs await_cleanup/signal_cleanup_done -- see CleanupWaiters below.
      @cleanup_waiters = CleanupWaiters.new
      ObjectSpace.define_finalizer(self, self.class.finalizer(socket))
    end

    # This connection's ONE dispatcher: the thread every callback on it runs
    # on, in arrival order, one at a time. Per connection rather than per
    # object because that is the promise the README makes to a caller who
    # shares state between an Application callback and a Workbook callback --
    # they can never be inside their blocks at the same time, so no lock of
    # their own is needed. See Dispatcher.
    #
    # This Client and its Dispatcher hold each other, and an attached Events
    # holds this Client back through the Dispatcher's target table. That ring
    # is collected as a ring: no thread ever holds a strong reference to any
    # of it across a park, so the whole connection is still collectible with
    # callbacks registered on it -- which is what the finalizer above needs
    # to be true if it is ever to close the socket.
    attr_reader :dispatcher

    # Register a consumer of server-initiated frames (those with no `id`).
    # `Events` is the consumer. The block is called INLINE on the reader
    # thread, which is the only thread reading this socket: the reader runs
    # nothing but the hand-off it is given here, and the hand-off must neither
    # block nor raise. A block that waited would stall every response on the
    # connection; one that made a COM call of its own would leave nobody to
    # read the answer and deadlock against itself. Enqueue and return.
    #
    # The block is also called once with `nil` when the stream ends, so a
    # consumer parked on a queue can finish instead of blocking on it forever
    # -- including when it registers after the stream has already ended, in
    # which case it is handed `nil` right here.
    #
    # APPENDS, never replaces. A replacing registration would silently switch
    # an earlier consumer off. The events feature puts up exactly one sink
    # here for the whole connection (Dispatcher#attach) and routes by handle
    # on the dispatcher thread, but nothing about this method assumes that:
    # anything else on the connection can register its own consumer and must
    # not be switched off by the events feature arming itself, or the other
    # way round.
    #
    # A sink lives exactly as long as this Client. The strong reference is
    # held HERE rather than in the Mailbox, because the running reader thread
    # pins the Mailbox: a sink almost always closes over the object that owns
    # it, and that object almost always holds this Client, so a strong
    # reference from the Mailbox would pin the Client too -- no Client could
    # be collected, its finalizer would never run and its socket would stay
    # open for the life of the process, which is the very leak the Mailbox was
    # extracted to fix. The Mailbox keeps a weak reference and can therefore
    # reach a sink without keeping it (or its Client) alive. Nothing is lost:
    # a collected Client has already had its socket closed by its finalizer,
    # so there are no further events to deliver.
    def on_event(&block)
      @event_sinks_mutex.synchronize { @event_sinks << block }
      @mailbox.on_event(&block)
      self
    end

    # The way back out, and the reason it exists: a sink registered for the
    # life of the connection is a consumer that cannot be dismantled. A
    # connection whose last callback has been removed would otherwise keep an
    # entry here (holding the Dispatcher, its target table and its parked
    # thread alive) and an entry in the Mailbox that the reader walks for
    # every frame -- so a connection would go on paying for events after
    # every object on it had stopped listening. Measured on the code before
    # any of this came down: 50 proxies that registered one callback and
    # removed it left 51 live threads and 50 sink entries.
    #
    # Identity, not equality: two consumers can be `==` without being the same
    # registration, and removing the wrong one silently stops a live consumer.
    def off_event(block)
      @event_sinks_mutex.synchronize { @event_sinks.reject! { |sink| sink.equal?(block) } }
      @mailbox.off_event(block)
      self
    end

    def call(method, params = {})
      response = @mailbox.request(method, params)
      # A sentinel, not a frame: the end of the stream is something that
      # happened here, and saying so in the shape of a wire error would make
      # it indistinguishable from one the bridge really sent -- a bridge that
      # ever reported a WineOLE::ProtocolError of its own would have had it
      # rewritten into a local one.
      raise ProtocolError, 'connection closed' if response.equal?(Mailbox::CLOSED)

      if response['error']
        klass = response['error']['class']
        # WineOLE::InstanceClosingError is the one remote error class this
        # client resolves to its own local class rather than wrapping in a
        # generic RemoteError -- so a caller can rescue it directly instead
        # of pattern-matching on RemoteError#remote_class.
        if klass == 'WineOLE::InstanceClosingError'
          raise InstanceClosingError, response['error']['message']
        end
        raise RemoteError.new(klass, response['error']['message'])
      end
      response['result']
    end

    def close
      @mailbox.close
    end

    # Blocks until the dispatcher finishes the $cleanup for `seq` (the
    # client closure, then the release_event that follows it). If the
    # caller IS the dispatcher thread -- `ole_release` called from inside a
    # callback -- do not wait: the $cleanup frame is queued behind the
    # current callback and will only run after it returns, so waiting here
    # would deadlock the dispatcher against itself.
    def await_cleanup(seq)
      return if on_dispatcher_thread?

      @cleanup_waiters.await(seq)
    end

    # Called by the dispatcher once it has finished delivering $cleanup
    # `seq` (Task 8), to release whatever thread is parked in `await_cleanup`
    # for it.
    def signal_cleanup_done(seq)
      @cleanup_waiters.signal(seq)
    end

    # Is the calling thread this connection's own dispatcher thread?
    def on_dispatcher_thread?
      @dispatcher.on_thread?(Thread.current)
    end

    # The socket, the waiter table and the event sinks -- everything the
    # reader thread shares with the calling threads.
    #
    # This is a separate object for a garbage-collection reason, not a
    # tidiness one. A block captures its whole binding, `self` included, so a
    # `Thread.new { read_loop }` written inside `Client#initialize` would make
    # the running reader thread a permanent GC root for the Client: no Client
    # could ever be collected, the ObjectSpace finalizer above would never
    # run, and its socket would stay open for the life of the process. The
    # thread is started from inside a Mailbox instead, so what it pins is this
    # object -- which holds no reference back to the Client, not even through
    # the event sinks: those it reaches only weakly (see Client#on_event).
    class Mailbox
      # Handed to a waiter when the stream ends. An object, not a frame:
      # nothing the bridge can send is `equal?` to it, so `Client#call` can
      # tell "this connection is over" from anything the bridge reported.
      CLOSED = Object.new.freeze

      def initialize(socket)
        @socket = socket
        @next_id = 0
        # Guards the socket WRITE and the bookkeeping below -- never a wait
        # for a response. Holding a lock across the round trip is what made a
        # COM call from inside an event callback impossible to even send while
        # the main thread was waiting.
        @mutex = Mutex.new
        @waiters = {}
        # WeakRefs, and deliberately so: the Client owns the sinks (see
        # Client#on_event). This list only lets the reader thread reach them.
        @event_sinks = []
        @closed = false
      end

      def start
        @reader = Thread.new { read_loop }
        @reader.abort_on_exception = false
        self
      end

      def on_event(&block)
        closed = @mutex.synchronize do
          # A sink registered on a dead connection is not registered at all:
          # the reader is gone, no frame can ever reach it, and leaving it in
          # the list would only let a `fail_all_waiters` still in flight hand
          # it a second end-of-stream.
          @event_sinks << WeakRef.new(block) unless @closed
          @closed
        end

        # Outside the mutex, because this runs user code on the CALLER's
        # thread: a consumer that closed the client from its end-of-stream
        # branch would otherwise deadlock on a lock this method still held.
        # Without this hand-off a consumer that attached after the bridge died
        # -- user code putting a handler on an existing proxy -- would never be
        # told the stream had ended, and its dispatcher thread would park on an
        # empty queue for the life of the process.
        deliver(block, nil) if closed
        self
      end

      # Drops one sink, and every dead weak reference met on the way past --
      # the reader only prunes those when it delivers, and a connection that
      # is quiet between registrations would otherwise keep them.
      def off_event(block)
        @mutex.synchronize do
          @event_sinks.reject! do |ref|
            sink = begin
              ref.__getobj__
            rescue WeakRef::RefError
              nil
            end
            sink.nil? || sink.equal?(block)
          end
        end
        self
      end

      # Sends one request and blocks this thread -- and only this thread --
      # until the reader routes the matching response back.
      def request(method, params)
        slot = Waiter.new
        id = nil
        @mutex.synchronize do
          raise ProtocolError, 'connection closed' if @closed

          id = (@next_id += 1)
          @waiters[id] = slot
          @socket.write(JSON.generate({id: id, method: method, params: params}) + "\n")
        end

        slot.take
      ensure
        @mutex.synchronize { @waiters.delete(id) } if id
      end

      def close
        @mutex.synchronize { @closed = true }
        @socket.close
        # A sink runs on the reader thread, so a `close` called from inside one
        # would have that thread join itself -- ThreadError, raised out of the
        # one method whose whole job is to shut the connection down cleanly.
        # The reader is on its way out anyway: the socket above is closed, so
        # its next read ends the loop.
        @reader.join(2) if @reader && @reader != Thread.current
      end

      private

      def read_loop
        while (line = @socket.gets)
          frame = begin
            JSON.parse(line)
          rescue JSON::ParserError
            next
          end

          # `null`, `123` and `[]` are all valid JSON and none of them is a
          # frame. Skipped exactly like an unparseable line: asking a
          # non-Hash for `key?` would raise NoMethodError, which is not in
          # this loop's rescue list, so one such line would kill the reader
          # and every later call on the connection with it.
          next unless frame.is_a?(Hash)

          if frame.key?('id')
            slot = @mutex.synchronize { @waiters[frame['id']] }
            slot&.fill(frame)
          else
            # Handed off, never run here.
            dispatch_to_sinks(frame)
          end
        end
      rescue IOError, Errno::EBADF, Errno::ECONNRESET
        nil
      ensure
        fail_all_waiters
      end

      # A waiter that is never woken waits forever, so EOF has to reach every
      # one of them.
      def fail_all_waiters
        pending = @mutex.synchronize do
          @closed = true
          @waiters.values.tap { @waiters.clear }
        end
        pending.each { |w| w.fill(CLOSED) }

        # Tell every event consumer the stream is over, so its dispatcher
        # thread can finish instead of blocking on an empty queue forever.
        # Without this each Events leaks a thread for the life of the process.
        dispatch_to_sinks(nil)
      end

      # The sinks are copied out under the mutex and called WITHOUT it. Holding
      # it across the dispatch would make a call issued from inside an event
      # consumer unable to even reach the wire -- the reader would be holding
      # the very lock that guards the socket write, on the very thread the
      # consumer runs on.
      def dispatch_to_sinks(frame)
        live_sinks.each { |sink| deliver(sink, frame) }
      end

      # The sinks that are still alive, resolved from their weak references,
      # dropping the ones whose Client has been collected.
      def live_sinks
        @mutex.synchronize do
          live = []
          @event_sinks.select! do |ref|
            sink = begin
              ref.__getobj__
            rescue WeakRef::RefError
              nil
            end
            live << sink if sink
            !sink.nil?
          end
          live
        end
      end

      # One misbehaving consumer must not take the connection down with it.
      # The sinks share a single reader thread, so an exception raised out of
      # one of them would end the read loop -- every other consumer on this
      # connection would stop seeing events, and every later call would fail
      # with "connection closed". Reported rather than swallowed: a sink that
      # raises is a bug in the sink.
      def deliver(sink, frame)
        sink.call(frame)
      rescue StandardError => e
        warn "wineole: event consumer raised #{e.class}: #{e.message}"
      end

      # One response, handed from the reader thread to the caller.
      class Waiter
        def initialize
          @mutex = Mutex.new
          @cond = ConditionVariable.new
          @value = nil
        end

        # First write wins. The reader fills a waiter and wakes its caller, but
        # the caller has not yet re-acquired the Mailbox mutex to delete itself
        # from the table; if the read loop ends in that window, the sweep would
        # otherwise overwrite the answer the bridge really sent with "connection
        # closed" and the caller would raise for a request that succeeded. It
        # also hardens the nil sentinel `take` waits on: nothing can take a
        # value back once it is there.
        def fill(value)
          @mutex.synchronize do
            return unless @value.nil?

            @value = value
            @cond.broadcast
          end
        end

        def take
          @mutex.synchronize do
            @cond.wait(@mutex) while @value.nil?
            @value
          end
        end
      end
    end

    # The await/signal handshake behind Client#await_cleanup and
    # #signal_cleanup_done, pulled out of Client so it can be unit-tested
    # without a live connection -- building a real Client for this would
    # open a socket, and the coordination itself has nothing to do with the
    # wire. One mutex guards a ConditionVariable per in-flight `seq` and the
    # set of `seq`s the dispatcher has already finished -- the same shape as
    # Mailbox::Waiter above, generalized from one outstanding key to many.
    class CleanupWaiters
      # How long `await` will wait for a `seq` that never gets signalled
      # before giving up and returning anyway. A caller stuck here forever
      # because a bridge or a dispatcher died mid-cleanup would be a worse
      # failure than one that eventually gets control back, even if the
      # instance's fate at that point is unknown.
      TIMEOUT = 30

      def initialize
        @mutex = Mutex.new
        @conds = {}
        @done = {}
      end

      # Blocks the calling thread until `signal(seq)` is called from another
      # thread, or TIMEOUT seconds pass, whichever comes first. Returns
      # immediately, without waiting at all, when `seq` was already
      # signalled before this call started -- the same "first write wins,
      # a late arrival still sees it" shape as Mailbox::Waiter.
      def await(seq)
        @mutex.synchronize do
          cond = (@conds[seq] ||= ConditionVariable.new)
          deadline = Process.clock_gettime(Process::CLOCK_MONOTONIC) + TIMEOUT
          until @done[seq]
            left = deadline - Process.clock_gettime(Process::CLOCK_MONOTONIC)
            break if left <= 0

            cond.wait(@mutex, left)
          end
          # Cleared on the way out so a `seq` (sequence numbers are not
          # reused) never accumulates an entry once nobody can still be
          # waiting on it.
          @done.delete(seq)
          @conds.delete(seq)
        end
      end

      def signal(seq)
        @mutex.synchronize do
          @done[seq] = true
          @conds[seq]&.broadcast
        end
      end
    end

    # Is the bridge on the other end of this connection reachable only through
    # the loopback interface -- i.e. the same machine?
    #
    # Deliberately the same test the bridge itself uses to decide whether a
    # token is required (`peer_addr().ip().is_loopback()` in main.rs). Anything
    # that keys off "is this local" -- path conversion in the Office wrapper,
    # for one -- must agree with the bridge, or a connection ends up remote for
    # authentication and local for everything else.
    #
    # `IPAddr#loopback?` covers all of 127.0.0.0/8 and ::1, matching Rust's
    # `IpAddr::is_loopback`. A host's own NIC address is NOT loopback, and that
    # is intended: it is a different machine as far as this boundary cares.
    def loopback?
      # The `false` pins peeraddr's reverse-lookup behaviour, which otherwise
      # follows `BasicSocket.do_not_reverse_lookup` -- a process-wide mutable
      # default owned by whatever application embeds this library, not by
      # this method. Index [3] is the numeric address regardless of the
      # setting, but without `false` a host app that flips that global
      # triggers an OS-level reverse DNS/PTR lookup as a side effect of
      # computing a tuple whose other elements are discarded here. That
      # lookup is slowest exactly where it matters: a loopback PTR resolves
      # instantly from /etc/hosts, while a remote peer's can hang on an
      # unreachable or misconfigured resolver -- blocking the very check
      # meant to tell loopback from non-loopback peers.
      IPAddr.new(@socket.peeraddr(false)[3]).loopback?
    rescue StandardError
      false
    end

    def create(class_name, cleanup: nil)
      Proxy.create(class_name, self, cleanup: cleanup)
    end

    def connect(class_name, cleanup: nil)
      Proxy.connect(class_name, self, cleanup: cleanup)
    end

    def connect_or_create(class_name, cleanup: nil)
      Proxy.connect_or_create(class_name, self, cleanup: cleanup)
    end
  end
end
