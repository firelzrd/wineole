require 'English'
require_relative '../client'

module WineOLE
  module MSOffice
    # Linux <-> Wine path conversion, and the question of whether converting
    # is meaningful at all.
    #
    # It is meaningful only when the client and the bridge are on one machine
    # looking at one Wine prefix. Convert when that is not true and you get a
    # path that silently refers to some other machine's filesystem.
    module Paths
      # Z:\..., C:/..., \\server\share\...
      WINDOWS_SHAPED = %r{\A(?:[A-Za-z]:[\\/]|\\\\)}.freeze

      # Deliberately keyed off the same loopback test the bridge uses to
      # decide whether a token is required. If the two definitions of "local"
      # drifted apart, a connection could be remote for authentication and
      # local for paths at once.
      #
      # A host's own NIC address counts as remote. Enumerating local
      # interfaces to notice otherwise would be more code, more edge cases
      # (containers, NAT, temporary IPv6 addresses), and would reintroduce
      # exactly that split.
      def self.convertible?(client:, windows: Client::WINDOWS)
        return false if windows # already Windows paths, and no winepath here

        client.loopback?
      end

      # Linux path -> Wine path. Returns the argument unchanged when it
      # already looks like a Windows path, and when winepath is unavailable
      # or fails.
      #
      # Failing to convert is not fatal -- the caller can write a Windows
      # path themselves, and will see that they need to. Raising here would
      # turn a recoverable inconvenience into a stopped script.
      def self.to_wine(path)
        return path if path.to_s.match?(WINDOWS_SHAPED)

        run_winepath('-w', path) || path
      end

      # Wine path -> Linux path. Same failure stance.
      def self.to_local(path)
        return path unless path.to_s.match?(WINDOWS_SHAPED)

        run_winepath('-u', path) || path
      end

      def self.run_winepath(flag, path)
        out = IO.popen(['winepath', flag, path.to_s], err: File::NULL, &:read)
        return nil unless $CHILD_STATUS&.success?

        out.strip.empty? ? nil : out.strip
      rescue SystemCallError, IOError
        nil
      end
      private_class_method :run_winepath
    end
  end
end
