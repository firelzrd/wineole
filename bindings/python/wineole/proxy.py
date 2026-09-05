import datetime

from .errors import NotSerializableError, StaleReferenceError
# Stated here rather than left to __init__.py: `ole_events` below names the
# class, so anything that imports this module alone -- tests/test_proxy.py
# does -- must get a working property rather than a NameError the first time
# it is used. events.py imports Proxy at call time, not import time, so there
# is no cycle.
from .events import Events


class Member:
    """A bound-but-not-yet-invoked member access, returned by `Proxy.__getattr__`.

    Python's attribute access (`proxy.Foo`) and a subsequent call (`(...)`)
    are two separate language-level operations -- unlike Ruby's
    `method_missing`, which treats them identically. Returning this wrapper
    from `__getattr__`, rather than performing the RPC immediately, keeps
    `proxy.Worksheets()` (property-get, zero args) and
    `proxy.Worksheets().Add(After=sheet)` (method call with a real Python
    keyword argument, mapping straight onto the wire's `named` dict)
    unambiguous -- at the cost of always needing the trailing `()`, even for
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
    def create(cls, class_name, client, cleanup=None):
        params = {'class_name': class_name}
        if cleanup is not None:
            params['cleanup'] = cls._build_cleanup(cleanup)
        handle = client.call('create', params)['$ole_ref']
        cls._register_cleanup(client, handle, cleanup)
        return cls(client, session_id=client.generation, handle=handle, created=True)

    @classmethod
    def connect(cls, class_name, client, cleanup=None):
        params = {'class_name': class_name}
        if cleanup is not None:
            params['cleanup'] = cls._build_cleanup(cleanup)
        handle = client.call('connect', params)['$ole_ref']
        cls._register_cleanup(client, handle, cleanup)
        return cls(client, session_id=client.generation, handle=handle, created=False)

    @classmethod
    def connect_or_create(cls, class_name, client, cleanup=None):
        params = {'class_name': class_name}
        if cleanup is not None:
            params['cleanup'] = cls._build_cleanup(cleanup)
        result = client.call('connect_or_create', params)
        handle = result['$ole_ref']
        cls._register_cleanup(client, handle, cleanup)
        return cls(client, session_id=client.generation, handle=handle, created=result['created'])

    @staticmethod
    def _build_cleanup(cleanup):
        # `cleanup` is {'steps': [[name, *args], ...], 'on_cleanup': callable
        # or absent}; the wire wants {'steps': [{'name':, 'args':}, ...],
        # 'callback': bool}. `callback` tells the bridge whether to hold the
        # root open and emit a $cleanup event for a registered closure, or
        # just run the steps and release outright -- so its value comes from
        # whether on_cleanup is present, not from anything the caller states
        # separately. The closure itself never goes on the wire.
        on_cleanup = cleanup.get('on_cleanup')
        if on_cleanup is not None and not callable(on_cleanup):
            raise TypeError('cleanup on_cleanup must be callable, or absent')
        steps = [{'name': step[0], 'args': list(step[1:])} for step in cleanup.get('steps', [])]
        return {'steps': steps, 'callback': on_cleanup is not None}

    @staticmethod
    def _register_cleanup(client, handle, cleanup):
        # Register the client closure (if any) so the dispatcher can deliver
        # $cleanup for this root handle. Only reached when the caller asked
        # for one, so a client that never uses on_cleanup never touches the
        # dispatcher here.
        if not cleanup:
            return
        on_cleanup = cleanup.get('on_cleanup')
        if on_cleanup is None:
            return
        client.dispatcher.register_cleanup(handle, on_cleanup)

    @classmethod
    def wrap(cls, client, session_id, ole_ref):
        return cls(client, session_id=session_id, handle=ole_ref, created=None)

    def __init__(self, client, session_id, handle, created):
        self._client = client
        self._session_id = session_id
        self._handle = handle
        self._created = created
        # Set here, and not merely left to the property below, because
        # __getattr__ answers any name it does not find with a Member: an
        # unset _ole_events would come back as a callable COM member stand-in
        # rather than as "not memoised yet".
        self._ole_events = None

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
        Proxy (e.g. xl.Worksheets()) -- attach-vs-create isn't a meaningful
        question for those."""
        return self._created

    @property
    def ole_events(self):
        """COM events for this object. Named with the ole_ prefix like every
        other bookkeeping member here: a Proxy forwards unknown names straight
        to COM, so a bare `events` would shadow a real Events member.

        A property, so it is found on the class before __getattr__ is
        consulted and never reaches COM. Memoised, because the Events owns a
        place on the connection dispatcher and a bridge-side subscription set:
        a fresh one per access would mean `on` and the `off` that is meant to
        undo it talked to different objects."""
        self._check_live()
        if self._ole_events is None:
            self._ole_events = Events(self._client, self._handle, self)
        return self._ole_events

    def __getattr__(self, name):
        # Only reached for names not already found by normal attribute
        # lookup (i.e. never for _client/_handle/_session_id/_created, which
        # __init__ sets via the real __dict__, nor for the ole_* properties
        # above) -- everything else is assumed to be a COM member name and
        # gets deferred into a Member.
        #
        # Dunders are the exception. CPython looks up *most* special methods
        # on the type, bypassing __getattr__ entirely -- which is why this
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
        result = self._client.call('release', {'handle': self._handle})
        # A client closure must run before this handle is actually gone: the
        # bridge answers with the $cleanup sequence number instead of
        # releasing outright, and this blocks until the dispatcher has
        # delivered it and the release_event that follows. See
        # Client.await_cleanup.
        if isinstance(result, dict):
            seq = result.get('cleanup')
            if seq is not None:
                self._client.await_cleanup(seq)
        return None

    def ole_leave_open(self):
        # Revokes the bridge's permission to run this instance's declared
        # cleanup steps when its last user releases the root -- the
        # instance (e.g. an auto-created Excel) then outlives this
        # connection instead of being torn down. Matches proxy.rb's
        # ole_leave_open.
        self._client.call('leave_open', {'handle': self._handle})
        return None

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
        if isinstance(value, datetime.datetime):
            # The same tag the receive side emits for VT_DATE. The wall clock
            # goes as-is: a VT_DATE carries no timezone, so converting here
            # would silently move the value the caller wrote.
            #
            # datetime is checked before date because datetime is a *subclass*
            # of date. Kept in this order deliberately: today both branches
            # format the same object with the same call, so the order is not
            # load-bearing (there is nothing observable to test). It is
            # defence against the future -- any date-specific formatting
            # added to the `date` branch below would silently truncate every
            # datetime without this order. Matching wineole/proxy.rb's
            # `encode`.
            return {'$type': 'time', 'iso8601': value.strftime('%Y-%m-%dT%H:%M:%S')}
        if isinstance(value, datetime.date):
            # A bare date, with no time component. It becomes midnight;
            # datetime was already handled above. Matching wineole/proxy.rb's
            # `encode`.
            return {'$type': 'time', 'iso8601': value.strftime('%Y-%m-%dT%H:%M:%S')}
        if isinstance(value, dict):
            return {k: self._encode(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [self._encode(v) for v in value]
        return value

    @classmethod
    def decode(cls, client, value):
        """One wire value, in Python terms.

        On the class rather than private to an instance because two different
        holders of a client need it: an invoke's RESULT, and an event's
        ARGUMENTS (Events._build_args). A second copy of this walk over there
        is how the same tagged value ends up a datetime when a call returns it
        and a raw {'$type': 'time'} dict when an event carries it -- and how a
        nested $ole_ref reaches a callback unwrapped.

        Recursive, mirroring proxy.rb's decode: a bulk range read comes back
        as a list of rows, and the values needing conversion sit inside it,
        not at the top level. A non-recursive decode would hand back raw
        {'$type': 'time'} dicts for every date cell in the range.
        """
        if isinstance(value, list):
            return [cls.decode(client, v) for v in value]
        if isinstance(value, dict):
            if '$ole_ref' in value:
                return cls.wrap(client, client.generation, value['$ole_ref'])
            if value.get('$type') == 'time':
                return datetime.datetime.fromisoformat(value['iso8601'])
            return {k: cls.decode(client, v) for k, v in value.items()}
        return value

    def _decode(self, value):
        return Proxy.decode(self._client, value)
