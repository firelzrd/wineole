class WineOLEError(Exception):
    pass


class NotSerializableError(WineOLEError):
    pass


class StaleReferenceError(WineOLEError):
    pass


class ProtocolError(WineOLEError):
    pass


class RemoteError(WineOLEError):
    def __init__(self, remote_class, message):
        self.remote_class = remote_class
        super().__init__(f"{remote_class}: {message}")
