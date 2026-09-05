import threading
import weakref

from .errors import warn


class Subscription:
    """One registration: a name, a callback and the args flag it asked for.

    Deliberately without __eq__, and not for tidiness: value equality would
    make two subscriptions with the same name, callback and flag equal without
    being the same registration -- registering one function twice for an event
    is enough. off(first) would then remove the second one too, unsubscribe,
    and leave the caller with a Subscription that never fires again and
    nothing said about it. A registration is the thing itself, so identity is
    its equality.

    `args` is per callback, but the wire has one flag per event: what goes on
    it is the union (see Events._effective_args).
    """

    __slots__ = ('_name', '_callback', '_args')

    def __init__(self, name, callback, args):
        self._name = name
        self._callback = callback
        self._args = args

    @property
    def name(self):
        return self._name

    @property
    def callback(self):
        return self._callback

    @property
    def args(self):
        return self._args

    def __repr__(self):
        return f"<wineole.Subscription {self._name!r} args={self._args!r}>"


class Events:
    """COM events for one object, reached as `proxy.ole_events`.

    Registering a callback is the only thing a caller does. Everything under
    it is DERIVED from that. The bridge-side subscription and the COM Advise
    beneath it belong to THIS object: they are put up by its first callback
    and taken back down by the last one removed from it. The sink on the
    connection and the dispatcher thread belong to the CONNECTION, so they
    follow the same rule one level up -- up with the first callback anywhere
    on the client, down with the last one anywhere on it (see Dispatcher).
    Making any of them a separate control would allow the state where a
    callback is registered and the event never arrives, with nothing to show
    for it.

    That invariant is a claim about ORDER as much as about bookkeeping, so
    on, off and close hold _wire across "decide, then tell the bridge".
    Deciding under one lock and writing the wire outside it lets two threads
    reach the bridge in the opposite order to the one they decided in: a
    subscribe landing after the unsubscribe that was meant to follow it leaves
    a registered callback whose event can never arrive, and the mirror case
    leaves the bridge advised with nothing to deliver to.

    Callbacks run on one dispatcher thread PER CONNECTION, in arrival order,
    one at a time -- so two objects on one client never run their callbacks
    concurrently, and a caller needs no lock between them. That thread is not
    here: it belongs to the connection, and this object attaches itself to it
    for as long as it has a registration.

    Mirrors WineOLE::Events in bindings/ruby/lib/wineole/events.rb.
    """

    def __init__(self, client, handle, proxy=None):
        self._client = client
        self._handle = handle
        # Weakly: the Proxy owns this object (Proxy.ole_events memoises it),
        # and the connection's Dispatcher holds this object strongly for as
        # long as it has a registration -- so a strong reference back would
        # keep every Proxy that ever registered a callback alive for the life
        # of the connection.
        self._proxy_ref = None if proxy is None else weakref.ref(proxy)
        # A plain dict, deliberately not a defaultdict: with a default
        # factory, merely LOOKING an event name up writes a permanent key --
        # an arriving event nobody subscribed to, or an off for a name that
        # was never on, would grow this table for the life of the connection.
        self._subs = {}
        self._lock = threading.Lock()
        # Held across the whole of on/off/close -- the decision AND the call
        # that carries it out -- and always acquired BEFORE _lock, never while
        # holding it. _lock is therefore never held across a wire round trip.
        # A callback calling off from the dispatcher waits for at most one
        # round trip and cannot deadlock: the thread it waits for needs
        # nothing from the dispatcher to finish.
        #
        # The connection's Dispatcher has a lock of its own, taken under this
        # one and never under _lock, so those two are never held together. The
        # Dispatcher takes none of these, and drops its own before it delivers,
        # which is what lets a callback register on another object.
        self._wire = threading.RLock()
        self._error_handler = None
        # Whether this object currently has a place on the connection's
        # dispatcher. Read and written under _wire only, which is what keeps
        # _disarm idempotent -- and it has to be: it is reached from off, from
        # close, and from on's failure path, on an object that may never have
        # been armed at all.
        self._attached = False

    @property
    def proxy(self):
        """The Proxy this belongs to, or None -- for an Events built directly
        against a client and handle, and for one whose Proxy has been
        collected."""
        return None if self._proxy_ref is None else self._proxy_ref()

    def on(self, name, callback, *, args=True):
        """Register `callback` for the COM event `name`.

        `args=False` tells the bridge not to mint handles for this event's
        object arguments. The callback is then called with no arguments at
        all. Worth it for a high-frequency event you only want to count.

        The bridge holds ONE flag per event, so what goes on the wire is the
        union: arguments are minted while any callback for that event wants
        them. Registering an args=True callback next to an args=False one
        therefore re-subscribes rather than leaving the first registration's
        flag standing -- measured on Excel before this was here: the second
        callback was handed nothing, having asked for the objects, and nothing
        said so.
        """
        if not callable(callback):
            raise TypeError('on needs a callback -- there is nothing to call otherwise')

        sub = Subscription(name, callback, args)
        with self._wire:
            # Before the subscribe, never after: the bridge advises the COM
            # source as the subscribe is handled, so an event can be on its
            # way back before the call returns. Arming afterwards would drop
            # it.
            self._arm()
            with self._lock:
                before = self._effective_args(name)
                self._subs.setdefault(name, []).append(sub)
                after = self._effective_args(name)
                wanted = None if after == before else after
            if wanted is not None:
                try:
                    self._client.call(
                        'subscribe', {'handle': self._handle, 'event': name, 'args': wanted})
                except Exception:
                    # A subscribe the bridge refused -- an object that is not
                    # an event source is the ordinary case -- must not leave
                    # the callback registered. Keeping it would produce
                    # exactly the state this class exists to make unreachable:
                    # a callback that is never called, with the error already
                    # raised and gone.
                    with self._lock:
                        self._drop(name, sub)
                        empty = not self._subs
                    if empty:
                        self._disarm()
                    raise
        return sub

    def off(self, name_or_subscription):
        """Take a callback back off, by the Subscription `on` returned (that
        registration and no other) or by event name (every callback for it).
        An unknown subscription or a name that was never registered is a
        no-op. Returns None."""
        with self._wire:
            sub = name_or_subscription if isinstance(name_or_subscription, Subscription) else None
            name = sub.name if sub is not None else name_or_subscription
            with self._lock:
                before = self._effective_args(name)
                self._drop(name, sub)
                after = self._effective_args(name)
                empty = not self._subs
            # Nothing was registered for this event, so nothing was derived
            # from it either. `after is None` alone cannot tell "the last
            # callback just went" from "there was never one".
            if before is None:
                return None

            try:
                if after is None:
                    # The last callback for this event is gone, so the
                    # subscription that only existed to feed it goes too --
                    # and with the last name for the object, the COM Advise
                    # underneath it.
                    self._client.call('unsubscribe', {'handle': self._handle, 'event': name})
                elif after != before:
                    # Callbacks remain, but the one that wanted arguments was
                    # among those removed: stop paying for handles nobody
                    # asked for.
                    self._client.call(
                        'subscribe', {'handle': self._handle, 'event': name, 'args': after})
            finally:
                # In a finally because the registry has already been emptied
                # above: if the unsubscribe raised (a connection that has just
                # gone is the ordinary case) the local half must still come
                # down, or this object keeps a dispatcher thread and a sink
                # entry on the client for the life of the connection with
                # nothing left to deliver.
                if empty:
                    self._disarm()
        return None

    def close(self):
        """Every callback forgotten, every subscription and Advise released,
        the dispatcher stopped and the sink taken off the connection.

        off-ing the last callback does all of this already -- that is the
        derivation this class is built on, and it is why there is no close you
        are REQUIRED to call. This is the bulk form, for a caller who does not
        want to remember which names it registered. `on` afterwards works
        exactly as the first one did: the object arms itself again."""
        with self._wire:
            with self._lock:
                names = list(self._subs)
            for name in names:
                with self._lock:
                    self._drop(name, None)
                try:
                    self._client.call('unsubscribe', {'handle': self._handle, 'event': name})
                except Exception:
                    # A connection that has already gone has already unadvised
                    # everything on it. Unlike off, this is not reported:
                    # close is what a caller reaches for when it is done, and
                    # the local half it exists to release comes down either
                    # way.
                    pass
            self._disarm()
        return None

    def on_error(self, handler):
        """ONE error handler per object, last writer wins -- deliberately not
        on's append. An error handler is not a subscription: it has no
        arguments to negotiate, nothing is derived from it, and there is
        nothing for a second one to add that the first cannot do.
        `on_error(None)` restores the default (a line on stderr). Returns this
        object, so it chains."""
        if handler is not None and not callable(handler):
            raise TypeError('on_error needs a callable, or None to restore the default')
        with self._lock:
            self._error_handler = handler
        return self

    # --- test-only helpers -------------------------------------------------
    # Production code never needs any of these, because callbacks are the
    # delivery mechanism. The queue and the thread the last two ask about
    # belong to the connection, so both are the Dispatcher's answers.

    def _registered_names(self):
        with self._lock:
            return list(self._subs)

    def _dispatcher_stopped(self, seconds):
        return self._client.dispatcher._stopped(seconds)

    def _drain(self, seconds=5):
        self._client.dispatcher._drain(seconds)
        return self

    # --- internals ---------------------------------------------------------

    def _arm(self):
        """Take this object's place on the connection: the frames for its
        handle start being routed to it, and the connection's dispatcher
        thread and sink go up if this is the first object to ask for them.
        Derived from there being a callback at all, so an ole_events a caller
        merely touched costs nothing anywhere. Called from `on` under _wire,
        which is what serializes it against _disarm."""
        if self._attached:
            return
        self._client.dispatcher.attach(self._handle, self)
        self._attached = True

    def _disarm(self):
        """The other half of _arm, once the last callback is gone: a
        registered callback is the only reason for any of this to exist.
        Without it an Events that has had every callback removed still holds a
        place on the connection's dispatcher and an entry the reader walks for
        every frame on it. The thread and the sink themselves only go with the
        LAST object to leave; that decision belongs to the Dispatcher, which
        is the only thing that knows whether any other object still has
        callbacks."""
        if not self._attached:
            return
        self._attached = False
        # By identity: the Dispatcher removes this object's routing entry and
        # nobody else's.
        self._client.dispatcher.detach(self._handle, self)

    def _effective_args(self, name):
        """What the wire flag for `name` should be: None when no callback is
        registered for it at all, otherwise True if any of them wants the
        event's object arguments minted. Call it under _lock."""
        subs = self._subs.get(name)
        if not subs:
            return None
        return any(sub.args for sub in subs)

    def _drop(self, name, sub):
        """Remove one subscription, or every callback for `name` when `sub` is
        None, and take the name out of the table when nothing is left for it.
        Call it under _lock.

        Identity, not equality: off removes the registration it was handed and
        no other. See Subscription."""
        subs = self._subs.get(name)
        if subs is None:
            return
        remaining = [] if sub is None else [s for s in subs if s is not sub]
        if remaining:
            self._subs[name] = remaining
        else:
            del self._subs[name]

    def _deliver(self, frame):
        """One frame, to the callbacks registered for its name. Called by the
        connection's Dispatcher, which routed it here by handle, and private
        for the reason the whole class exists: a registered callback is the
        only way in, and a public deliver would say otherwise. The frame's
        argument handles are given back by the Dispatcher, in a finally that
        covers a frame this raised on as well as one that reached nobody at
        all.

        Never raises."""
        with self._lock:
            subs = list(self._subs.get(frame.get('event'), ()))
        try:
            args = self._build_args(frame.get('args'))
        except Exception as exc:
            # A frame the bridge cannot have sent. Reported once, and no
            # callback is called with it -- the Dispatcher still releases its
            # handles.
            self._report(exc, frame)
            return
        for sub in subs:
            try:
                sub.callback(*args)
            except BaseException as exc:  # noqa: BLE001
                # Everything, for the reason the dispatcher catches
                # everything: that thread is the whole delivery mechanism, and
                # a callback raising something outside the usual hierarchy
                # must not take the next callback, the next event and every
                # later release down with it.
                self._report(exc, frame)

    def _build_args(self, raw):
        if raw is None:
            return []
        if not isinstance(raw, list):
            # Checked explicitly, unlike Ruby's equivalent, because a Python
            # str is iterable: without this a frame whose args is a string
            # would call the callback once per character instead of being
            # reported as the malformed frame it is.
            raise TypeError(f"event args must be a list or null, got {type(raw).__name__}")
        # The same decode an invoke's result goes through, so an event
        # argument that is an object arrives as a Proxy and one that is a date
        # arrives as a datetime.
        #
        # Imported here rather than at the top of the module: proxy.py imports
        # this module so that `ole_events` works for anything that loads it
        # alone, and a matching import up there would be a cycle. By the time
        # a frame is delivered the module is long since loaded, so this is a
        # dict lookup.
        from .proxy import Proxy

        return [Proxy.decode(self._client, value) for value in raw]

    def _report(self, exc, frame):
        """Never raises. It is called from the dispatcher's own except, so an
        exception out of here is the one thing that could still end the
        thread."""
        try:
            with self._lock:
                handler = self._error_handler
            event = frame.get('event') if isinstance(frame, dict) else None
            if handler is None:
                warn(f"wineole: {event} callback raised {type(exc).__name__}: {exc}")
                return
            try:
                handler(exc, frame)
            except BaseException as inner:  # noqa: BLE001
                # An on_error that itself raises must not recurse, and must
                # not be able to kill the dispatcher either.
                warn(f"wineole: on_error raised {type(inner).__name__} "
                     f"while reporting {type(exc).__name__}")
        except BaseException:  # noqa: BLE001
            pass  # even stderr being gone must not end the dispatcher
