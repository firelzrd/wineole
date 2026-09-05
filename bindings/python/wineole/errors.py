import sys


class WineOLEError(Exception):
    pass


class NotSerializableError(WineOLEError):
    pass


class StaleReferenceError(WineOLEError):
    pass


class ProtocolError(WineOLEError):
    pass


class InstanceClosingError(WineOLEError):
    """Raised by a call that arrived after the bridge decided this instance's
    root proxy is on its way out (its declared cleanup steps are running, or
    have already run). Distinguished from a generic RemoteError so a caller
    can catch "this instance is closing" specifically rather than pattern-
    matching on RemoteError.remote_class. Mirrors wineole/errors.rb's
    WineOLE::InstanceClosingError."""
    pass


class RemoteError(WineOLEError):
    def __init__(self, remote_class, message):
        self.remote_class = remote_class
        super().__init__(f"{remote_class}: {message}")


def warn(message):
    """Write one line to stderr about something this library must not raise.

    The reader thread, the dispatcher thread and Events.on_error all have the
    same problem: they are the whole delivery mechanism for a connection, so
    an exception escaping any of them is permanent and silent. They report
    instead -- through here, in one place, so the three of them cannot drift
    into three different formats. `sys.stderr` is looked up at call time, so a
    test wrapping a block in contextlib.redirect_stderr captures it.

    Never raises: even stderr being gone must not end a dispatcher.
    """
    try:
        print(message, file=sys.stderr)
    except Exception:
        pass
