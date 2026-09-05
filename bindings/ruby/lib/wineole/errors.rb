module WineOLE
  class Error < StandardError; end
  class NotSerializableError < Error; end
  class StaleReferenceError < Error; end
  class ProtocolError < Error; end
  # Raised by a call that arrived after the bridge decided this instance's
  # root proxy is on its way out (a client `$cleanup` closure is running, or
  # has already run). Distinguished from a generic RemoteError so a caller
  # can rescue "this instance is closing" specifically rather than pattern-
  # matching on RemoteError#remote_class.
  class InstanceClosingError < Error; end

  class RemoteError < Error
    attr_reader :remote_class

    def initialize(remote_class, message)
      @remote_class = remote_class
      super("#{remote_class}: #{message}")
    end
  end
end
