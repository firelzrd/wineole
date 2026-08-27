module WineOLE
  class Error < StandardError; end
  class NotSerializableError < Error; end
  class StaleReferenceError < Error; end
  class ProtocolError < Error; end

  class RemoteError < Error
    attr_reader :remote_class

    def initialize(remote_class, message)
      @remote_class = remote_class
      super("#{remote_class}: #{message}")
    end
  end
end
