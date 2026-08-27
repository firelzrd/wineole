"""Python client for wineole — mirrors Ruby's `lib/wineole.rb` front door.

Re-exports the same public names the Ruby entry point does, so that
`import wineole; wineole.Client(...)` works without callers needing to know
the internal module layout.
"""

import threading

from .errors import WineOLEError, NotSerializableError, StaleReferenceError, ProtocolError, RemoteError
from .client import Client
from .proxy import Proxy, Member

__all__ = [
    'WineOLEError', 'NotSerializableError', 'StaleReferenceError', 'ProtocolError', 'RemoteError',
    'Client', 'Proxy', 'Member',
    'open', 'create', 'connect', 'connect_or_create', 'close',
]

_default_client = None
_default_client_lock = threading.Lock()


def open(**kwargs):
    """Open a client the same way Client.open does, and also make it the
    package's implicit default -- a subsequent create()/connect() call uses
    this client rather than lazily creating a separate zero-config one.

    Named `open` deliberately, matching Client.open -- this shadows the
    `open` builtin only via `from wineole import *` (discouraged anyway,
    and not something this package's own code does); `wineole.open(...)`
    itself never collides with anything.

    Calling open() again replaces the implicit default without closing
    whatever it previously pointed to -- the caller owns the returned
    client and is responsible for closing it if not relying on the
    implicit default's eventual __del__-based cleanup."""
    global _default_client
    client = Client.open(**kwargs)
    with _default_client_lock:
        _default_client = client
    return client


def _get_default_client():
    """The lazily-initialized implicit default client used by create()/
    connect() when nothing was ever explicitly opened via open(). Thread-safe:
    the fast path avoids locking once initialized; only the first caller(s)
    racing on a None default contend for the lock, and only the actual
    winner calls Client.open -- everyone else sees it already set once they
    acquire the lock and never calls Client.open again."""
    global _default_client
    if _default_client is not None:
        return _default_client
    with _default_client_lock:
        if _default_client is None:
            _default_client = Client.open()
        return _default_client


def create(class_name):
    return _get_default_client().create(class_name)


def connect(class_name):
    return _get_default_client().connect(class_name)


def connect_or_create(class_name):
    return _get_default_client().connect_or_create(class_name)


def close():
    """Closes the implicit default client, if one exists, and clears it so
    the next create()/connect() lazily opens a fresh one. Mainly for test
    hygiene; also usable to release the implicit default early in a
    long-running process."""
    global _default_client
    with _default_client_lock:
        if _default_client is not None:
            _default_client.close()
        _default_client = None
