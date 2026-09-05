class Passthrough:
    """Shared by every wrapper around a COM object that otherwise passes
    unknown names straight through to it: the `ole` property, plus
    `__getattr__`/`__setattr__`.

    Deliberately does not define `__init__` -- the including classes take
    different constructor arguments (`Range(proxy)` vs.
    `Sheet(proxy, version)` vs. `Book(proxy, client, version,
    convert_paths=True)`), so each sets `_ole` itself, first thing.

    Two consequences, both deliberate:

    * Every wrapper-owned attribute is underscore-prefixed, because
      `__setattr__` sends anything else to COM. `xl.Visible = True` reaches
      COM exactly as it does on a bare `Proxy`.
    * `__getattr__` is consulted only for names normal lookup did not find,
      so a wrapper method wins over a COM member of the same spelling.
      That is why the wrapper's names were checked against live Excel for
      collisions on the Ruby side and are kept identical here (`close`
      deliberately shadows `Workbook.Close`; see Book).
    """

    @property
    def ole(self):
        """The underlying Proxy, for reaching COM explicitly."""
        return self._ole

    def _passthrough_target(self):
        """Where unknown names go. The sole override hook: a wrapper that
        forwards somewhere other than the object it holds (a control that
        forwards to the thing with Caption and Value rather than to the
        OLEObject host around it) overrides this and nothing else."""
        return self._ole

    def __getattr__(self, name):
        # Only reached for names not already found by normal attribute
        # lookup -- never for `_ole` (set in the subclass __init__ via the
        # real __dict__), never for `ole` or for any method defined here or
        # on a subclass.
        #
        # Dunders are the exception, for the same reason Proxy.__getattr__
        # refuses them: CPython looks up most special methods on the type,
        # but some stdlib protocols probe the *instance* (copy.deepcopy
        # does a plain getattr(obj, '__deepcopy__', None)). Forwarding
        # those turns a should-be-AttributeError into a real RPC for a
        # member name COM can never have. COM member names are never Python
        # dunders, so refusing them here cannot collide with anything real.
        if name.startswith('__') and name.endswith('__'):
            raise AttributeError(name)
        if name == '_ole':
            # A half-built instance (__init__ raised, copy.copy, pickle)
            # has no _ole yet; without this the lookup recurses forever.
            raise AttributeError(name)
        return getattr(self._passthrough_target(), name)

    def __setattr__(self, name, value):
        if name.startswith('_'):
            object.__setattr__(self, name, value)
        else:
            # A COM property set. Proxy.__setattr__ turns this into an
            # invoke of `name=`, which is why the wrapper does not need to
            # know anything about the wire here.
            setattr(self._passthrough_target(), name, value)

    def __iter__(self):
        # Without this, iter()/for/`in` fall back to the legacy 0-based
        # __getitem__ sequence protocol -- and on Excel and Sheet that
        # __getitem__ is the addressing DSL, so `for x in sheet` would fire
        # it with 0, 1, 2 and hand COM nonsense. Proxy refuses iteration
        # for the same reason; refuse it the same way.
        raise TypeError(f"{type(self).__name__} is not iterable")
