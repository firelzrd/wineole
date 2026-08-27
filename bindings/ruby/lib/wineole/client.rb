require 'socket'
require 'json'
require 'rbconfig'
require 'tmpdir'
require_relative 'errors'
require_relative 'proxy'

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
      TCPSocket.new(host, port)
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
      @next_id = 0
      @mutex = Mutex.new
      ObjectSpace.define_finalizer(self, self.class.finalizer(socket))
    end

    def call(method, params = {})
      @mutex.synchronize do
        id = (@next_id += 1)
        @socket.write(JSON.generate({id: id, method: method, params: params}) + "\n")
        line = @socket.gets
        raise ProtocolError, 'connection closed' if line.nil?
        response = JSON.parse(line)
        unless response['id'] == id
          raise ProtocolError, "id mismatch: expected #{id}, got #{response['id']}"
        end
        raise RemoteError.new(response['error']['class'], response['error']['message']) if response['error']
        response['result']
      end
    end

    def close
      @socket.close
    end

    def create(class_name)
      Proxy.create(class_name, self)
    end

    def connect(class_name)
      Proxy.connect(class_name, self)
    end

    def connect_or_create(class_name)
      Proxy.connect_or_create(class_name, self)
    end
  end
end
