import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
sys.path.insert(0, os.path.dirname(__file__))

from support.fake_com import FakeComObject

from wineole.msoffice.passthrough import Passthrough


class Wrapper(Passthrough):
    """A minimal concrete user of the mixin, the way Range/Sheet/Book/Excel
    each are: it sets `_ole` first thing and adds a couple of names of its
    own."""

    def __init__(self, ole):
        self._ole = ole

    def close(self, save=False):
        # A deliberate shadow, exactly as Book#close is: COM has a Close
        # member, and this name must win over it.
        return ('wrapper close', save)


class Redirected(Wrapper):
    """The one override hook: forwards somewhere other than `_ole`."""

    def __init__(self, ole, other):
        Wrapper.__init__(self, ole)
        self._other = other

    def _passthrough_target(self):
        return self._other


class MSOfficePassthroughTest(unittest.TestCase):
    def test_ole_exposes_the_underlying_object(self):
        ole = FakeComObject()
        self.assertIs(ole, Wrapper(ole).ole)

    def test_an_unknown_name_reaches_com_and_needs_the_call_parentheses(self):
        w = Wrapper(FakeComObject(Name='Sheet1'))
        # The trailing () is the whole Python/Ruby difference: a COM member
        # access hands back something callable, and the wrapper must not
        # swallow that.
        self.assertEqual('Sheet1', w.Name())

    def test_a_wrapper_method_wins_over_a_com_member_of_the_same_name(self):
        w = Wrapper(FakeComObject(close='the com member'))
        self.assertEqual(('wrapper close', False), w.close(),
                         '__getattr__ is only consulted when normal lookup fails, '
                         'so a wrapper method must shadow the COM member')

    def test_a_property_set_reaches_com(self):
        ole = FakeComObject()
        w = Wrapper(ole)
        w.Visible = True
        self.assertEqual({'Visible': True}, ole.writes)

    def test_an_underscore_prefixed_attribute_stays_on_the_wrapper(self):
        ole = FakeComObject()
        w = Wrapper(ole)
        w._version = '11.0'
        self.assertEqual('11.0', w._version)
        self.assertEqual({}, ole.writes,
                         'every wrapper-owned attribute is underscore-prefixed, '
                         'and must never become a COM property set')

    def test_a_dunder_is_not_forwarded_to_com(self):
        w = Wrapper(FakeComObject())
        # copy.deepcopy and friends probe the instance with a plain getattr;
        # answering those with a forwarded call turns a should-be
        # AttributeError into a real round trip for a name COM cannot have.
        with self.assertRaises(AttributeError):
            w.__deepcopy__

    def test_the_wrapper_is_not_iterable(self):
        w = Wrapper(FakeComObject())
        with self.assertRaises(TypeError) as ctx:
            iter(w)
        self.assertIn('not iterable', str(ctx.exception))

    def test_the_passthrough_target_is_the_override_hook(self):
        host = FakeComObject(Name='the host')
        other = FakeComObject(Name='the target')
        r = Redirected(host, other)
        self.assertEqual('the target', r.Name())
        r.Caption = 'x'
        self.assertEqual({'Caption': 'x'}, other.writes)
        self.assertEqual({}, host.writes)
        self.assertIs(host, r.ole, 'ole still names the object this wrapper holds')

    def test_a_half_built_wrapper_raises_attribute_error_not_recursion(self):
        # __new__ without __init__ is what copy.copy and pickle produce, and
        # what a debugger sees when __init__ raised before setting _ole.
        w = Wrapper.__new__(Wrapper)
        with self.assertRaises(AttributeError):
            w.Name
        with self.assertRaises(AttributeError):
            w._ole


if __name__ == '__main__':
    unittest.main()
