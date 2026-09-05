require_relative 'wineole/errors'
require_relative 'wineole/client'
require_relative 'wineole/proxy'
require_relative 'wineole/dispatcher'
require_relative 'wineole/events'

module WineOLE
  @default_client = nil
  @mutex = Mutex.new

  # Opens a client the same way Client.open does, and also makes it the
  # module's implicit default -- a subsequent WineOLE.create/.connect call
  # uses this client rather than lazily creating a separate zero-config one.
  # Calling .open again replaces the implicit default without closing
  # whatever it previously pointed to -- the caller owns the returned
  # client and is responsible for closing it if not relying on the
  # implicit default's eventual GC-finalizer cleanup.
  def self.open(...)
    client = Client.open(...)
    @mutex.synchronize { @default_client = client }
    client
  end

  # The lazily-initialized implicit default client used by .create/.connect
  # when nothing was ever explicitly opened via WineOLE.open. Thread-safe:
  # the fast path avoids locking once initialized; only the first caller(s)
  # racing on a nil default contend for the lock, and only the actual
  # winner calls Client.open -- everyone else sees it already set once they
  # acquire the lock, and the `||=` means they never call Client.open again.
  #
  # Public (a deliberate exception to "the core layer is not changed" --
  # the same exception, and for the same reason, as Client#loopback?):
  # bundled wrappers such as WineOLE::MSOffice::Excel need the Client their
  # Proxy belongs to (Book needs it to answer #loopback?), and this module is
  # already the layer holding that connection. Making the caller guess, or
  # open a second connection just to answer that question, would be worse
  # than letting the layer that already knows the answer say so.
  def self.default_client
    return @default_client if @default_client

    @mutex.synchronize { @default_client ||= Client.open }
  end

  def self.create(class_name)
    default_client.create(class_name)
  end

  def self.connect(class_name)
    default_client.connect(class_name)
  end

  def self.connect_or_create(class_name)
    default_client.connect_or_create(class_name)
  end

  # Closes the implicit default client, if one exists, and clears it so the
  # next .create/.connect lazily opens a fresh one. Mainly for test hygiene
  # (clearing state between test cases); also usable to release the
  # implicit default early in a long-running process.
  def self.close
    @mutex.synchronize do
      @default_client&.close
      @default_client = nil
    end
  end
end
