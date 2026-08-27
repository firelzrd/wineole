import datetime

from .errors import NotSerializableError, StaleReferenceError


class Member:
    """A bound-but-not-yet-invoked member access, returned by `Proxy.__getattr__`.

    Python's attribute access (`proxy.Foo`) and a subsequent call (`(...)`)
    are two separate language-level operations — unlike Ruby's
    `method_missing`, which treats them identically. Returning this wrapper
    from `__getattr__`, rather than performing the RPC immediately, keeps
    `proxy.Worksheets()` (property-get, zero args) and
    `proxy.Worksheets().Add(After=sheet)` (method call with a real Python
    keyword argument, mapping straight onto the wire's `named` dict)
    unambiguous — at the cost of always needing the trailing `()`, even for
    pure properties.
    """

    __slots__ = ('_proxy', '_name')

    def __init__(self, proxy, name):
        object.__setattr__(self, '_proxy', proxy)
        object.__setattr__(self, '_name', name)

    def __call__(self, *args, **kwargs):
        return self._proxy.invoke(self._name, list(args), kwargs)


class Proxy:
    @classmethod
    def create(cls, class_name, client):
        handle = client.call('create', {'class_name': class_name})['$ole_ref']
        return cls(client, session_id=client.generation, handle=handle, created=True)

    @classmethod
    def connect(cls, class_name, client):
        handle = client.call('connect', {'class_name': class_name})['$ole_ref']
        return cls(client, session_id=client.generation, handle=handle, created=False)

    @classmethod
    def connect_or_create(cls, class_name, client):
        result = client.call('connect_or_create', {'class_name': class_name})
        return cls(client, session_id=client.generation, handle=result['$ole_ref'], created=result['created'])

    @classmethod
    def wrap(cls, client, session_id, ole_ref):
        return cls(client, session_id=session_id, handle=ole_ref, created=None)

    def __init__(self, client, session_id, handle, created):
        self._client = client
        self._session_id = session_id
        self._handle = handle
        self._created = created

    @property
    def ole_handle(self):
        return self._handle

    @property
    def ole_session_id(self):
        return self._session_id

    @property
    def ole_created(self):
        """Was this instance freshly created by connect_or_create's
        fallback, or attached to something already running? True for
        .create, False for .connect, whatever the bridge reported for
        .connect_or_create, and None for anything derived from another
        Proxy (e.g. xl.Worksheets()) — attach-vs-create isn't a meaningful
        question for those."""
        return self._created

    def __getattr__(self, name):
        # Only reached for names not already found by normal attribute
        # lookup (i.e. never for _client/_handle/_session_id/_created, which
        # __init__ sets via the real __dict__, nor for the ole_* properties
        # above) — everything else is assumed to be a COM member name and
        # gets deferred into a Member.
        #
        # Dunders are the exception. CPython looks up *most* special methods
        # on the type, bypassing __getattr__ entirely — which is why this
        # class needs no Ruby-style implicit-conversion guard list. But some
        # stdlib protocols probe the *instance*: copy.deepcopy() does a
        # plain getattr(obj, '__deepcopy__', None). Answering those with a
        # Member turns a should-be-AttributeError into a real RPC for a
        # member name COM can never have (DISP_E_UNKNOWNNAME), and hides
        # __reduce__'s NotSerializableError from copy/pickle. COM member
        # names are never Python dunders, so refusing them here cannot
        # collide with anything real.
        if name.startswith('__') and name.endswith('__'):
            raise AttributeError(name)
        return Member(self, name)

    def __iter__(self):
        # Without this, iter()/for/`in` fall back to the legacy 0-based
        # __getitem__ sequence protocol, firing one real RPC round-trip per
        # index (invoke name='' args=[0], [1], [2], ...) until the remote
        # errors out. Fail fast instead.
        raise TypeError(f"{type(self).__name__} is not iterable")

    def __setattr__(self, name, value):
        if name.startswith('_'):
            object.__setattr__(self, name, value)
        else:
            self.invoke(name + '=', [value], {})

    def __getitem__(self, key):
        self._check_live()
        args = list(key) if isinstance(key, tuple) else [key]
        return self.invoke('', args, {})

    def ole_const_load(self):
        self._check_live()
        return self._client.call('const_load', {'handle': self._handle})

    def ole_release(self):
        return self._client.call('release', {'handle': self._handle})

    def __reduce__(self):
        raise NotSerializableError(
            'Proxy references are connection-scoped and cannot be persisted'
        )

    def invoke(self, name, args, named):
        # Deliberately bare and public, unlike every other meta-method here
        # (which are ole_-prefixed to avoid shadowing a same-named remote
        # COM member): an explicit escape hatch for the rare case a COM
        # object really does define e.g. an `ole_handle` member, matching
        # real Ruby WIN32OLE's own choice to keep `invoke` public and
        # unprefixed.
        self._check_live()
        params = {
            'handle': self._handle,
            'name': name,
            'args': [self._encode(a) for a in args],
            'named': {k: self._encode(v) for k, v in named.items()},
        }
        return self._decode(self._client.call('invoke', params))

    def _check_live(self):
        if self._session_id != self._client.generation:
            raise StaleReferenceError('this reference belongs to a previous connection')

    def _encode(self, value):
        if isinstance(value, Proxy):
            # The argument's own liveness is not enough: a Proxy belonging
            # to a *different* Client is live from its own point of view,
            # but its handle id means nothing (or something unrelated) in
            # the receiver's connection. Check it against the receiver's
            # client, matching wineole/proxy.rb's `encode`.
            if value.ole_session_id != self._client.generation:
                raise StaleReferenceError(
                    'this reference belongs to a different connection and cannot be '
                    'passed as an argument here'
                )
            return {'$ole_ref': value.ole_handle}
        if isinstance(value, dict):
            return {k: self._encode(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [self._encode(v) for v in value]
        return value

    def _decode(self, value):
        if isinstance(value, dict) and '$ole_ref' in value:
            return Proxy.wrap(self._client, self._client.generation, value['$ole_ref'])
        if isinstance(value, dict) and value.get('$type') == 'time':
            return datetime.datetime.fromisoformat(value['iso8601'])
        return value
