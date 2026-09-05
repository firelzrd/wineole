require_relative '../proxy'

module WineOLE
  module MSOffice
    # Shared by every wrapper around a COM object that otherwise passes
    # unknown methods straight through to it: the `ole` reader, plus
    # `method_missing`/`respond_to_missing?`. `Range`, `Sheet`, `Book`,
    # `Control` and `Form` each need exactly these three members; `Control`
    # overrides `passthrough_target`; extracted here so they are written
    # once instead of copied three times.
    #
    # Deliberately does not define `initialize` -- the including classes
    # take different constructor arguments (`Range.new(proxy)` vs.
    # `Sheet.new(proxy, version:)` vs. `Book.new(proxy, client:, version:,
    # convert_paths:)`), so each sets `@ole` itself.
    module Passthrough
      # The underlying Proxy, for reaching COM explicitly.
      attr_reader :ole

      def method_missing(name, *args, &block)
        return super if Proxy::IMPLICIT_CONVERSIONS.include?(name)

        passthrough_target.public_send(name, *args, &block)
      end

      def respond_to_missing?(name, include_private = false)
        # Anything not named here is a COM member as far as this class is
        # concerned -- the same stance Proxy takes, and for the same reason:
        # Ruby probes to_ary, to_str, coerce and friends behind the scenes,
        # and answering `true` to those makes e.g. `puts obj` try a
        # conversion that ends in NoMethodError from inside the interpreter.
        # Reuse Proxy's list rather than keeping a second copy of it.
        !Proxy::IMPLICIT_CONVERSIONS.include?(name)
      end

      private

      # Where unknown methods go. `@ole` for every wrapper except `Control`,
      # which forwards to the object that has Caption and Value rather than
      # to the OLEObject host around it (see Control's class comment).
      def passthrough_target
        @ole
      end
    end
  end
end
