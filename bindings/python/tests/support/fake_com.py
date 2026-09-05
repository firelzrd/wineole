"""Hand-written stand-ins for COM objects, spelled the way the Python raw
client spells them.

The one rule every fake here follows: a COM member is a CALLABLE, and
reading a property and calling a method are the same operation
(`range.Font()`, `borders.Item(7)`, `rows.Count()`). That is exactly what
`wineole.proxy.Member` does -- `__getattr__` hands back something that has
to be called, because Python cannot tell a property read from a method call
the way Ruby's `method_missing` can. A property SET is a plain
`setattr(obj, 'Bold', True)`, which `wineole.proxy.Proxy.__setattr__` turns
into an `invoke('Bold=', ...)`; the fakes record those in `writes`.

Shared by every msoffice unit test file so the eight of them do not each
redefine a FakeRange that has drifted from the others.
"""
import re

from wineole.errors import RemoteError
from wineole.msoffice.paths import Paths


class _Recorder:
    """Base for a fake whose property sets are what a test asserts on.

    `__setattr__` is overridden, so the constructor has to reach past it
    with `object.__setattr__` for the fake's own bookkeeping attributes --
    otherwise `self.writes = {}` would be recorded as a COM write of a
    member called `writes` and there would be no dict to record it into.
    """

    def __init__(self):
        object.__setattr__(self, 'writes', {})

    def __setattr__(self, name, value):
        self.writes[name] = value


class FakeCount:
    """What `Rows()` and `Columns()` answer: an object whose Count is a call."""

    def __init__(self, n):
        self._n = n

    def Count(self):
        return self._n


class FakeComObject:
    """A COM object with a fixed set of readable members.

    `FakeComObject(Name='Sheet1')` answers `obj.Name()` with `'Sheet1'` and
    raises AttributeError for anything else, so a test that expects a
    passthrough to reach a member proves the member was actually asked for.
    """

    def __init__(self, **members):
        object.__setattr__(self, 'writes', {})
        object.__setattr__(self, '_members', dict(members))

    def __getattr__(self, name):
        # Dunders are refused for the same reason Proxy refuses them: some
        # stdlib protocols probe the instance, and answering them with a
        # callable turns a should-be-AttributeError into a fake COM call.
        if name.startswith('__') and name.endswith('__'):
            raise AttributeError(name)
        if name not in self._members:
            raise AttributeError(f"fake COM object has no member {name!r}")
        value = self._members[name]
        return lambda *args, **kwargs: value

    def __setattr__(self, name, value):
        self.writes[name] = value


class FakeFont(_Recorder):
    pass


class FakeInterior(_Recorder):
    pass


class FakeBorder(_Recorder):
    pass


class FakeBorders(_Recorder):
    """The Borders collection: a bulk assignment target AND an Item() source.

    `fetches` counts Item() calls, so a test can prove the bulk path never
    took the per-edge loop (and the other way round).
    """

    def __init__(self):
        _Recorder.__init__(self)
        object.__setattr__(self, 'items', {})
        object.__setattr__(self, 'fetches', 0)

    def Item(self, index):
        object.__setattr__(self, 'fetches', self.fetches + 1)
        if index not in self.items:
            self.items[index] = FakeBorder()
        return self.items[index]


# Whatever this Excel calls the General format. Deliberately not the value
# this project's own host returns: a real spelling here would bake one
# machine's locale into the suite and would let an implementation that
# hardcodes that spelling pass.
LOCAL_GENERAL = 'LOCAL-GENERAL-FORMAT-NAME'


class FakeApplication:
    def __init__(self):
        self.international_calls = []

    def International(self, index):
        self.international_calls.append(index)
        return LOCAL_GENERAL


class FakeComRange(_Recorder):
    """A COM Range: enough for Range's shape checks and its passthrough.

    `written` is what the last `Value =` assignment carried -- the exact
    shape that would have crossed the wire. `Value()` keeps answering
    whatever the range was constructed with, because Excel's own Value is
    not what a write changed from this object's point of view; the tests
    that care read `written`.
    """

    def __init__(self, rows=1, cols=1, value=None):
        _Recorder.__init__(self)
        object.__setattr__(self, '_rows', rows)
        object.__setattr__(self, '_cols', cols)
        object.__setattr__(self, '_value', value)
        object.__setattr__(self, 'written', None)
        object.__setattr__(self, 'interior', FakeInterior())

    def Rows(self):
        return FakeCount(self._rows)

    def Columns(self):
        return FakeCount(self._cols)

    def Value(self):
        return self._value

    def Interior(self):
        return self.interior

    def __setattr__(self, name, value):
        self.writes[name] = value
        if name == 'Value':
            object.__setattr__(self, 'written', value)


class FakeComRangeForFormat(_Recorder):
    """A COM Range as Format sees it.

    Font and Interior are separate objects, as they are in COM, and each
    counts how many times it was fetched so the round-trip saving can be
    asserted rather than assumed.
    """

    def __init__(self):
        _Recorder.__init__(self)
        object.__setattr__(self, 'font', FakeFont())
        object.__setattr__(self, 'interior', FakeInterior())
        object.__setattr__(self, 'application', FakeApplication())
        object.__setattr__(self, 'borders', FakeBorders())
        object.__setattr__(self, 'font_fetches', 0)
        object.__setattr__(self, 'interior_fetches', 0)
        object.__setattr__(self, 'borders_fetches', 0)

    def Font(self):
        object.__setattr__(self, 'font_fetches', self.font_fetches + 1)
        return self.font

    def Interior(self):
        object.__setattr__(self, 'interior_fetches', self.interior_fetches + 1)
        return self.interior

    def Application(self):
        return self.application

    def Borders(self):
        object.__setattr__(self, 'borders_fetches', self.borders_fetches + 1)
        return self.borders


class SelectSpy:
    """Mixed into every fake COM object in the Excel test file.

    Any accidental use of `Application.Range` (which requires selecting a
    sheet first) shows up as a recorded call rather than as a silently
    passing test. Reset in setUp, asserted empty in tearDown -- so every
    single test in that file doubles as a check that `Excel[...]` never
    calls Select.
    """

    calls = []

    @classmethod
    def reset(cls):
        SelectSpy.calls = []

    def Select(self):
        SelectSpy.calls.append(f"{type(self).__name__}.Select")


class FakeComWorksheet(SelectSpy):
    """A COM Worksheet. Records what Range/Cells were asked for so a test
    can assert the right cells were addressed without a real Excel."""

    def __init__(self, name='Sheet1'):
        self.name = name
        self.range_calls = []
        self.cells_calls = []
        self.ranges = []

    def Name(self):
        return self.name

    # Measured against a live Excel 11: Worksheet.CodeName is the module's
    # own name, independent of the visible tab Name.
    def CodeName(self):
        return self.name

    def Range(self, addr):
        self.range_calls.append(addr)
        r = FakeComRange()
        self.ranges.append(r)
        return r

    def Cells(self, row, col):
        self.cells_calls.append([row, col])
        r = FakeComRange()
        self.ranges.append(r)
        return r


class FakeComWorksheets:
    """A COM Worksheets collection -- Application.Worksheets and
    Workbook.Worksheets are both exercised through this one fake."""

    def __init__(self, items):
        self.items = list(items)
        self.add_after_calls = []
        self.item_calls = []

    def Item(self, name_or_index):
        self.item_calls.append(name_or_index)
        if isinstance(name_or_index, int) and not isinstance(name_or_index, bool):
            return self.items[name_or_index - 1]
        for w in self.items:
            if w.name == name_or_index:
                return w
        raise KeyError(f"no worksheet named {name_or_index!r}")

    def Count(self):
        return len(self.items)

    def Add(self, After=None):
        # `After` is spelled exactly as it goes on the wire: a real Python
        # keyword argument, which Proxy turns into COM's `named` dict.
        self.add_after_calls.append(After)
        new_sheet = FakeComWorksheet(f"Sheet{len(self.items) + 1}")
        self.items.append(new_sheet)
        return new_sheet


class FakeComWorkbook(SelectSpy):
    def __init__(self, name='Book1', worksheets=None, path='', full_name='',
                 vb_project=None, vb_project_denied=False):
        self.name = name
        self.worksheets = FakeComWorksheets(worksheets if worksheets is not None else [])
        self.active_sheet = None
        self.path = path
        self.full_name = full_name
        self.save_as_calls = []
        self.close_calls = []
        self.vb_project = vb_project
        self.vb_project_denied = vb_project_denied
        # Measured against a live Excel 11: a worksheet's CodeName is ""
        # until something has touched this workbook's VBProject. This flag
        # is how FakeComWorksheetWithProject models that.
        self.vb_project_touched = False

    def VBProject(self):
        if self.vb_project_denied:
            raise RemoteError('WIN32OLERuntimeError', 'COM error (0x800A03EC)')
        if self.vb_project is None:
            self.vb_project = FakeVBProject()
        self.vb_project_touched = True
        return self.vb_project

    def Worksheets(self):
        return self.worksheets

    def ActiveSheet(self):
        return self.active_sheet

    def Path(self):
        return self.path

    def FullName(self):
        return self.full_name

    def SaveAs(self, path):
        self.save_as_calls.append(path)
        return 'saved'

    def Close(self, save_changes):
        self.close_calls.append(save_changes)
        return 'closed'


class FakeComWorkbooks:
    def __init__(self, items):
        self.items = list(items)
        self.add_calls = 0

    def Item(self, name):
        for w in self.items:
            if w.name == name:
                return w
        raise KeyError(f"no workbook named {name!r}")

    def Add(self):
        self.add_calls += 1
        wb = FakeComWorkbook(f"Book{len(self.items) + 1}",
                             worksheets=[FakeComWorksheet('Sheet1')])
        self.items.append(wb)
        return wb


class FakeComApplication(SelectSpy):
    """Stands in for the COM Excel.Application AND the Proxy that wraps it
    -- Excel takes whatever it is handed as `_ole` with no distinction.

    `ole_created` mirrors Proxy's own meaning and is a plain attribute, not
    a call, because Proxy exposes it as a property. `ole_release` and
    `ole_leave_open` are the two Proxy bookkeeping calls Excel delegates to;
    whether a release actually quits Excel is the bridge's decision, made
    from the `cleanup=` a real Client.create was handed, so this fake only
    records that the call happened.

    The four `fail_*` / `gone` knobs stand in for the ways the real
    Application can refuse mid-block. Ruby's test file reopened the
    singleton to do this; a flag reads the same and does not leak into the
    next test.
    """

    def __init__(self, created=True, version='11.0', workbooks=None, worksheets=None,
                 active_workbook=None, active_sheet=None,
                 display_alerts=True, screen_updating=True):
        self.ole_created = created
        self.version = version
        self.workbooks = FakeComWorkbooks(workbooks if workbooks is not None else [])
        self.worksheets = FakeComWorksheets(worksheets if worksheets is not None else [])
        self.active_workbook = active_workbook
        self.active_sheet = active_sheet
        self.quit_calls = 0
        self.display_alerts = display_alerts
        self.screen_updating = screen_updating
        self.display_alerts_history = []
        self.screen_updating_history = []
        self.ole_release_calls = 0
        self.ole_leave_open_calls = 0
        self.visible = None
        self.gone = False
        self.fail_display_alerts_read = False
        self.fail_screen_updating_read = False
        self.fail_display_alerts_restore = False

    # --- Proxy bookkeeping -------------------------------------------

    def ole_release(self):
        self.ole_release_calls += 1
        return None

    def ole_leave_open(self):
        self.ole_leave_open_calls += 1
        return None

    # --- COM members: reads are calls ---------------------------------

    def Version(self):
        return self.version

    def Workbooks(self):
        return self.workbooks

    def Worksheets(self):
        return self.worksheets

    def ActiveWorkbook(self):
        return self.active_workbook

    def ActiveSheet(self):
        return self.active_sheet

    def Quit(self):
        self.quit_calls += 1
        self.gone = True

    def DisplayAlerts(self):
        if self.gone:
            raise RuntimeError('COM error (0x800706BE)')
        if self.fail_display_alerts_read:
            raise RuntimeError('reading DisplayAlerts exploded')
        return self.display_alerts

    def ScreenUpdating(self):
        if self.gone:
            raise RuntimeError('COM error (0x800706BE)')
        if self.fail_screen_updating_read:
            raise RuntimeError('reading ScreenUpdating exploded')
        return self.screen_updating

    # --- COM members: writes are setattr ------------------------------

    def __setattr__(self, name, value):
        if name == 'Visible':
            object.__setattr__(self, 'visible', value)
        elif name == 'DisplayAlerts':
            if self.gone:
                raise RuntimeError('COM error (0x800706BE)')
            # The restore, not the suppression, is what fails: the fake lets
            # `False` through and dies only on the way back, which is the
            # order the real thing fails in.
            if self.fail_display_alerts_restore and value is not False:
                raise RuntimeError('COM error (0x800706BE)')
            self.display_alerts_history.append(value)
            object.__setattr__(self, 'display_alerts', value)
        elif name == 'ScreenUpdating':
            if self.gone:
                raise RuntimeError('COM error (0x800706BE)')
            self.screen_updating_history.append(value)
            object.__setattr__(self, 'screen_updating', value)
        else:
            object.__setattr__(self, name, value)


class FakeClient:
    """A stand-in for wineole.Client: `loopback` (a plain attribute here,
    a property on the real one) plus the three factory methods, each
    recording the class name and the `cleanup=` it was handed."""

    def __init__(self, app=None, loopback=True):
        self.app = app
        self.loopback = loopback
        self.create_calls = []
        self.connect_calls = []
        self.connect_or_create_calls = []
        self.create_cleanups = []
        self.connect_cleanups = []
        self.connect_or_create_cleanups = []

    def create(self, class_name, cleanup=None):
        self.create_calls.append(class_name)
        self.create_cleanups.append(cleanup)
        return self.app

    def connect(self, class_name, cleanup=None):
        self.connect_calls.append(class_name)
        self.connect_cleanups.append(cleanup)
        return self.app

    def connect_or_create(self, class_name, cleanup=None):
        self.connect_or_create_calls.append(class_name)
        self.connect_or_create_cleanups.append(cleanup)
        return self.app


def split_com_lines(text):
    """Split a CodeModule's text the way Ruby's String#split does.

    Python's re.split keeps a trailing empty field; Ruby's split drops every
    one of them. The block mechanism indexes into these lines and hands the
    leftovers back to a caller that compares them for equality, so the two
    languages have to agree on the length. Mirrors the production helper of
    the same name in wineole/msoffice/vba_block.py -- kept in sync by hand
    so the fake cannot flatter an implementation that diverged.
    """
    parts = re.split(r'\r?\n', text)
    while parts and parts[-1] == '':
        parts.pop()
    return parts


class FakeCodeModule:
    """Stands in for COM's CodeModule.

    Holds text the way Excel does -- reads always come back with CRLF,
    whatever was written -- and records how many times the whole body was
    fetched, so the round-trip count can be asserted.

    `lines=` bypasses the text split entirely, so a fake can hold blank
    lines directly: splitting "\\r\\n" yields no lines at all, not the two
    blank lines Excel actually reports for a module emptied of its blocks.
    """

    def __init__(self, text='', lines=None):
        if lines is not None:
            self.lines = list(lines)
        elif text == '':
            self.lines = []
        else:
            self.lines = split_com_lines(text)
        self.reads = 0

    def CountOfLines(self):
        return len(self.lines)

    def Lines(self, start, count):
        self.reads += 1
        if count == 0:
            raise AssertionError('Lines(1, 0) must never be called')
        chunk = self.lines[start - 1:start - 1 + count]
        return '\r\n'.join(chunk) + '\r\n'

    def AddFromString(self, text):
        # Excel inserts at the top, not the end.
        self.lines = split_com_lines(text) + self.lines

    def DeleteLines(self, start, count):
        del self.lines[start - 1:start - 1 + count]

    @property
    def text(self):
        return '\n'.join(self.lines)


class FakeVBComponent:
    """A VBComponent. `Name` and `Type` are reads, so they are calls;
    renaming is a property SET, so it is a plain setattr -- which is exactly
    how the wrapper spells it, and exactly where Excel's own rename can
    fail.

    Excel's Add-then-rename is not atomic: adding under a taken name
    succeeds and only the rename fails, leaving a stray behind. The fake has
    to be able to reach that state or the cleanup cannot be tested, which is
    what `rename_fails` is for.
    """

    def __init__(self, name, component_type=1):
        object.__setattr__(self, 'name', name)
        object.__setattr__(self, 'component_type', component_type)
        object.__setattr__(self, 'rename_fails', False)
        object.__setattr__(self, 'code_module', FakeCodeModule())
        object.__setattr__(self, 'export_bytes',
                           ('Attribute VB_Name = "%s"\r\n' % name).encode('ascii'))

    def Name(self):
        return self.name

    # Type mirrors VBComponents' own: 1 standard, 2 class, 3 UserForm, 100 a
    # module Excel owns (a worksheet's, ThisWorkbook's).
    def Type(self):
        return self.component_type

    def CodeModule(self):
        return self.code_module

    def Export(self, path):
        """Excel writes the file itself; the fake just puts the bytes there.

        vba_api hands this a Windows-shaped path (Paths.to_wine) -- real
        Wine resolves that back to the real file transparently at the OS
        layer, which this reproduces with Paths.to_local so the fake can
        still do plain Python file I/O on it. A no-op when the host has no
        winepath to have converted the path in the first place.
        """
        with open(Paths.to_local(path), 'wb') as handle:
            handle.write(self.export_bytes)

    def __setattr__(self, name, value):
        if name == 'Name':
            if self.rename_fails:
                raise RemoteError('WIN32OLERuntimeError', 'COM error (0x80020009)')
            object.__setattr__(self, 'name', value)
        else:
            object.__setattr__(self, name, value)


class FakeVBComponents:
    def __init__(self, project):
        self.project = project

    def Count(self):
        return len(self.project.component_list)

    def Item(self, name):
        for component in self.project.component_list:
            if component.name == name:
                return component
        raise RemoteError('X', 'not found')

    def Add(self, component_type):
        component = FakeVBComponent(
            'Module%d' % (len(self.project.component_list) + 1),
            component_type=component_type)
        component.rename_fails = self.project.next_rename_fails
        self.project.next_rename_fails = False
        self.project.component_list.append(component)
        return component

    def Remove(self, component):
        self.project.component_list.remove(component)

    def Import(self, path):
        with open(Paths.to_local(path), 'rb') as handle:
            self.project.imported_bytes = handle.read()
        component = FakeVBComponent('Imported')
        self.project.component_list.append(component)
        return component


class FakeVBProject:
    def __init__(self):
        self.component_list = []
        self.imported_bytes = None
        self.next_rename_fails = False
        self.vb_components = FakeVBComponents(self)

    def VBComponents(self):
        return self.vb_components

    @property
    def components(self):
        """A dict view keyed by each component's *current* Name, built fresh
        on every read. It cannot be a dict populated at Add() time: Add()
        and the rename are two separate steps, and a dict key does not
        follow a rename -- Excel's own VBComponents.Item looks up by
        whatever the Name currently is."""
        return {component.name: component for component in self.component_list}

    def add_existing(self, name, component_type=1):
        component = FakeVBComponent(name, component_type=component_type)
        self.component_list.append(component)
        return component


class FakeProjectWhoseLookupFails:
    """A project that opens fine but whose component lookup fails. Needed to
    tell "AccessVBOM is off" apart from "that component is not there": a
    rescue covering both would answer every one of them with registry
    advice."""

    def VBComponents(self):
        return self

    def Item(self, name):
        raise RemoteError('WIN32OLERuntimeError', 'COM error (0x800A0009)')


class FakeComWorksheetWithProject(FakeComWorksheet):
    """A worksheet whose Parent is a workbook with a VBProject.

    CodeName models the measured Excel 11 behaviour that the plain
    FakeComWorksheet does not: it is "" until something has touched the
    parent workbook's VBProject, and the sheet's real code name from then
    on. A SheetVBA that read the name before opening the project would hand
    VBComponents.Item("") and get 0x800A0009 -- this is the fake that makes
    that a failing test rather than a live-Excel surprise.
    """

    def __init__(self, name='Sheet1', vb_project=None, denied=False):
        FakeComWorksheet.__init__(self, name)
        self.parent = FakeComWorkbook(vb_project=vb_project,
                                      vb_project_denied=denied)

    def CodeName(self):
        return self.name if self.parent.vb_project_touched else ''

    def Parent(self):
        return self.parent


# --- controls and forms ---------------------------------------------------
#
# One rule these fakes add to the file's own: a fake COM object must NOT
# expose a real Python attribute named like a COM member it is supposed to
# refuse. `setattr(obj, 'Bogus', 1)` on a plain object succeeds in silence,
# which would let an implementation that never checks anything pass -- so
# every fake below whose job is to refuse a property put overrides
# __setattr__ and raises for the names it does not know.


def remote_error(hresult):
    """The RemoteError shape the bridge raises, with the HRESULT the Ruby
    fakes use: 0x800A03EC for "Excel refused this call" and 0x80020006 for
    "no such member"."""
    return RemoteError('WIN32OLERuntimeError', 'COM error (%s)' % hresult)


class FakeComRangeWithBox:
    """A COM Range with a position -- what Sheet[] hands geometry() through
    `.ole`. Reads are calls, as everywhere in this file."""

    def Left(self):
        return 10.5

    def Top(self):
        return 20.0

    def Width(self):
        return 30.0

    def Height(self):
        return 40.0


class FakeComSheetForGeometry:
    """Just enough Worksheet for Sheet[] to resolve an address, recording
    what was asked for so a test can prove Excel was NOT touched."""

    def __init__(self):
        self.range_calls = []

    def Name(self):
        return 'Sheet1'

    def Range(self, addr):
        self.range_calls.append(addr)
        return FakeComRangeWithBox()


class FakeItemHost:
    """A COM collection whose Item(name) knows exactly one name. Everything
    else raises, which is the "free" answer check_free is reading."""

    def __init__(self, known):
        self._known = known

    def Item(self, name):
        if name != self._known:
            raise remote_error('0x800A03EC')
        return 'found'


class FakeEventsRecorder:
    """Records subscriptions the way Events does, without a bridge. The
    signature is the Python one -- `on(name, callback, *, args=True)` -- so a
    Control that passed the callback as a keyword would fail here."""

    def __init__(self):
        self.on_calls = []
        self.off_calls = []
        self.closed = 0

    def on(self, name, callback, *, args=True):
        self.on_calls.append([name, callback, args])
        return 'subscription'

    def off(self, name_or_subscription):
        self.off_calls.append(name_or_subscription)
        return None

    def close(self):
        self.closed += 1
        return None


class FakeControlComObject:
    """A COM object that records property puts and hands out one Events.

    `Object` is a real COM member here, not bookkeeping: on an MSForms
    control it answers the raw control under the extender, and the wrapper's
    reader is deliberately called `ole_object` so that `ctl.Object()` still
    reaches this.
    """

    def __init__(self):
        object.__setattr__(self, 'puts_seen', [])
        object.__setattr__(self, 'ole_events', FakeEventsRecorder())

    def Object(self):
        return 'the-com-member'

    def __setattr__(self, name, value):
        self.puts_seen.append([name, value])


class FakeVBAWriter:
    """Records what a VBA writer was asked to write, and where. The keyword
    names are BookVBA.write's own, so a Control that spelled `into` some
    other way would fail here."""

    def __init__(self):
        self.writes = []

    def write(self, code, *, name='main', into=None):
        self.writes.append({'code': code, 'name': name, 'into': into})
        return self


class FakeFormForControl:
    """The Form a 'userform' Control talks to: hands out one runtime control
    per name, memoised, and records every request."""

    def __init__(self, name):
        self.name = name
        self.runtime_requests = []
        self._runtime = {}

    def runtime_control(self, control_name):
        self.runtime_requests.append(control_name)
        if control_name not in self._runtime:
            self._runtime[control_name] = FakeControlComObject()
        return self._runtime[control_name]


# One name registry (a dict name -> shape) backs Shapes, the nine legacy
# collections and OLEObjects, the way one sheet does in Excel: a control
# added through any collection is visible to Shapes.Item and Shapes.Count.
# That is what lets the rollback tests tell "added, then deleted" from
# "never added".
#
# The collection -> FormControlType map is spelled out here rather than
# imported from wineole.msoffice.controls on purpose: a fake that read the
# production table could not catch a wrong entry in it.
FORM_CONTROL_TYPE_BY_COLLECTION = {
    'Buttons': 0,
    'CheckBoxes': 1,
    'DropDowns': 2,
    'GroupBoxes': 4,
    'Labels': 5,
    'ListBoxes': 6,
    'OptionButtons': 7,
    'ScrollBars': 8,
    'Spinners': 9,
}


class FakeFormControlShape:
    """A legacy form control: what Buttons.Add / CheckBoxes.Add hand back.

    Excel refuses some renames (an invalid or reserved name) with 0x800A03EC
    and any unknown property with 0x80020006; this does both, so the rollback
    path has something real to fail on.
    """

    RECORDED_PUTS = ('Caption', 'OnAction')

    def __init__(self, registry, collection, box, name):
        object.__setattr__(self, 'registry', registry)
        object.__setattr__(self, 'collection', collection)
        object.__setattr__(self, 'box', box)
        object.__setattr__(self, 'name', name)
        object.__setattr__(self, 'puts_seen', [])
        object.__setattr__(self, 'deleted', False)
        registry[name] = self

    def Name(self):
        return self.name

    def FormControlType(self):
        return FORM_CONTROL_TYPE_BY_COLLECTION[self.collection]

    def Delete(self):
        self.registry.pop(self.name, None)
        object.__setattr__(self, 'deleted', True)

    def __setattr__(self, name, value):
        if name == 'Name':
            if value == 'Refused':
                raise remote_error('0x800A03EC')
            self.registry.pop(self.name, None)
            object.__setattr__(self, 'name', value)
            self.registry[value] = self
            self.puts_seen.append(['Name', value])
            return
        if name not in self.RECORDED_PUTS:
            raise remote_error('0x80020006')
        self.puts_seen.append([name, value])


class FakeLegacyCollection:
    """One of the nine: Buttons, CheckBoxes, and so on."""

    def __init__(self, registry, collection, add_calls):
        self.registry = registry
        self.collection = collection
        self.add_calls = add_calls

    def Add(self, left, top, width, height):
        self.add_calls.append([self.collection, left, top, width, height])
        return FakeFormControlShape(
            self.registry, self.collection, [left, top, width, height],
            '%s %d' % (self.collection, len(self.registry) + 1))

    def Item(self, name):
        shape = self.registry.get(name)
        if (not isinstance(shape, FakeFormControlShape)
                or shape.collection != self.collection):
            raise remote_error('0x800A03EC')
        return shape


class FakeMSFormsObject:
    """The MSForms control inside an OLEObject."""

    def __init__(self, progid):
        object.__setattr__(self, 'progid', progid)
        object.__setattr__(self, 'puts_seen', [])
        object.__setattr__(self, 'ole_events', FakeEventsRecorder())

    def __setattr__(self, name, value):
        # An Image has no Caption; Excel refuses the put.
        if self.progid == 'Forms.Image.1' and name == 'Caption':
            raise remote_error('0x80020006')
        self.puts_seen.append([name, value])


class FakeOLEObject:
    """An OLEObject host. `progid` is the attribute this fake keeps its own
    state in; `progID()` is the COM member Excel actually answers, spelled
    with Excel's own capitalisation."""

    HOST_PUTS = ('LinkedCell', 'ListFillRange', 'Visible', 'PrintObject',
                 'Placement')

    def __init__(self, registry, progid, box, name):
        object.__setattr__(self, 'registry', registry)
        object.__setattr__(self, 'progid', progid)
        object.__setattr__(self, 'box', box)
        object.__setattr__(self, 'name', name)
        object.__setattr__(self, 'puts_seen', [])
        object.__setattr__(self, 'deleted', False)
        object.__setattr__(self, 'inner', FakeMSFormsObject(progid))
        registry[name] = self

    def Name(self):
        return self.name

    def progID(self):
        return self.progid

    def Object(self):
        return self.inner

    def FormControlType(self):
        # Shape.FormControlType on an ActiveX shape is a COM error in Excel.
        raise remote_error('0x800A03EC')

    def Delete(self):
        self.registry.pop(self.name, None)
        object.__setattr__(self, 'deleted', True)

    def __setattr__(self, name, value):
        if name == 'Name':
            if value == 'Refused':
                raise remote_error('0x800A03EC')
            self.registry.pop(self.name, None)
            object.__setattr__(self, 'name', value)
            self.registry[value] = self
            self.puts_seen.append(['Name', value])
            return
        if name not in self.HOST_PUTS:
            raise remote_error('0x80020006')
        self.puts_seen.append([name, value])


class FakeOLEObjects:
    """The passthrough trap, reproduced: Excel 11 fails OLEObjects.Add with
    0x800A03EC unless all four of Left/Top/Width/Height come with ClassType.
    The keyword names are the ones that go on the wire, so an implementation
    that passed them positionally would fail here."""

    REQUIRED = ('ClassType', 'Left', 'Top', 'Width', 'Height')

    def __init__(self, registry, add_calls):
        self.registry = registry
        self.add_calls = add_calls

    def Add(self, **named):
        self.add_calls.append(dict(named))
        if any(key not in named for key in self.REQUIRED):
            raise remote_error('0x800A03EC')
        box = [named['Left'], named['Top'], named['Width'], named['Height']]
        return FakeOLEObject(self.registry, named['ClassType'], box,
                             'OLE%d' % (len(self.registry) + 1))

    def Item(self, name):
        shape = self.registry.get(name)
        if not isinstance(shape, FakeOLEObject):
            raise remote_error('0x800A03EC')
        return shape


class FakeShapesForControls:
    """Shapes sees every family, which is what makes check_free work across
    them."""

    def __init__(self, registry):
        self.registry = registry

    def Count(self):
        return len(self.registry)

    def Item(self, name):
        shape = self.registry.get(name)
        if shape is None:
            raise remote_error('0x800A03EC')
        return shape


class FakeComSheetForControls:
    """A COM Worksheet backing all three collections off one registry, with a
    VBProject behind it so both writers (the parent workbook's BookVBA and
    the sheet's own SheetVBA) can be exercised.

    `denied=True` makes VBProject refuse, which is how the AccessDenied path
    is reached.
    """

    def __init__(self, denied=False):
        self.registry = {}
        self.add_calls = []
        self.ole_add_calls = []
        self.range_calls = []
        self.project = FakeVBProject()
        # A worksheet's own code module, which Excel owns (Type 100).
        self.project.add_existing('Sheet1', 100)
        self.parent = FakeComWorkbook(name='Book1', vb_project=self.project,
                                      vb_project_denied=denied)

    def Name(self):
        return 'Sheet1'

    def CodeName(self):
        # Measured against Excel 11 and modelled by FakeComWorksheetWithProject
        # too: "" until something has touched the parent's VBProject.
        return 'Sheet1' if self.parent.vb_project_touched else ''

    def Parent(self):
        return self.parent

    def Shapes(self):
        return FakeShapesForControls(self.registry)

    def OLEObjects(self):
        return FakeOLEObjects(self.registry, self.ole_add_calls)

    def Range(self, addr):
        self.range_calls.append(addr)
        return FakeComRangeWithBox()

    def _legacy(self, collection):
        return FakeLegacyCollection(self.registry, collection, self.add_calls)

    def Buttons(self):
        return self._legacy('Buttons')

    def CheckBoxes(self):
        return self._legacy('CheckBoxes')

    def OptionButtons(self):
        return self._legacy('OptionButtons')

    def ListBoxes(self):
        return self._legacy('ListBoxes')

    def DropDowns(self):
        return self._legacy('DropDowns')

    def Spinners(self):
        return self._legacy('Spinners')

    def ScrollBars(self):
        return self._legacy('ScrollBars')

    def Labels(self):
        return self._legacy('Labels')

    def GroupBoxes(self):
        return self._legacy('GroupBoxes')


class FakeRuntimeControl:
    """A control on the LOADED form: fires events, has live values."""

    def __init__(self):
        object.__setattr__(self, 'ole_events', FakeEventsRecorder())
        object.__setattr__(self, 'puts_seen', [])

    def __setattr__(self, name, value):
        self.puts_seen.append([name, value])


class FakeRuntimeControls:
    """The loaded form's Controls collection: one control per name, made on
    demand and handed back the same afterwards."""

    def __init__(self):
        self._controls = {}

    def Item(self, name):
        if name not in self._controls:
            self._controls[name] = FakeRuntimeControl()
        return self._controls[name]


class FakeUserFormInstance:
    """What WineOLE_Form_<Name> hands back: the form's default instance."""

    def __init__(self):
        self.ole_events = FakeEventsRecorder()
        self.controls = FakeRuntimeControls()
        self.calls = []
        self.visible = False

    def Controls(self):
        return self.controls

    def Visible(self):
        return self.visible

    def Hide(self):
        # The extender's own Hide answers GetIDsOfNames and then refuses every
        # out-of-process Invoke, which is why the wrapper goes through VBA.
        # Recorded so a test can prove it was never called.
        self.calls.append('Hide')

    def mark_shown(self):
        self.visible = True


class FakeDesignControl:
    """A design-time MSForms control on the Designer."""

    def __init__(self, name):
        object.__setattr__(self, 'name', name)
        object.__setattr__(self, 'puts_seen', [])

    def Name(self):
        return self.name

    def __setattr__(self, name, value):
        if name == 'Bogus':
            raise remote_error('0x80020006')
        self.puts_seen.append([name, value])


class FakeDesignerControls:
    """Add takes (progid, name) -- the name at Add time, no rename step --
    and Remove takes the NAME, not the object. That asymmetry with
    Shapes/OLEObjects is real Excel, not a simplification."""

    def __init__(self):
        self.controls = {}
        self.add_calls = []

    def Add(self, progid, name):
        self.add_calls.append([progid, name])
        self.controls[name] = FakeDesignControl(name)
        return self.controls[name]

    def Item(self, name):
        control = self.controls.get(name)
        if control is None:
            raise remote_error('0x800A03EC')
        return control

    def Remove(self, name):
        return self.controls.pop(name, None)

    def Count(self):
        return len(self.controls)


class FakeDesigner:
    def __init__(self):
        self.controls = FakeDesignerControls()

    def Controls(self):
        return self.controls

    def Width(self):
        # An ordinary Designer member, for the passthrough test.
        return 240


class FakeVBComponentForForms:
    def __init__(self, name, component_type):
        self.name = name
        self.component_type = component_type
        self.designer = FakeDesigner()

    def Name(self):
        return self.name

    def Type(self):
        return self.component_type

    def Designer(self):
        return self.designer


class FakeVBComponentsForForms:
    def __init__(self, components):
        self.components = components

    def Item(self, name):
        component = self.components.get(name)
        if component is None:
            raise remote_error('0x800A0009')
        return component


class FakeVBProjectForForms:
    def __init__(self, components):
        self.components = components

    def VBComponents(self):
        return FakeVBComponentsForForms(self.components)


class FakeBookVBAForForms:
    """The BookVBA a Form talks to, as a recorder.

    `fail_write_with` lets a test make the helper write fail the way a
    refused Excel call would, so the rollback in Forms.add and the
    deliberately UN-rescued write in Forms[] can both be exercised.
    """

    def __init__(self):
        self.writes = []
        self.components = {'Module1': FakeVBComponentForForms('Module1', 1)}
        self.removed = []
        self.fail_write_with = None

    def write(self, code, *, name='main', into=None):
        if self.fail_write_with is not None:
            raise self.fail_write_with
        self.writes.append({'code': code, 'name': name, 'into': into})
        return self

    def add_component(self, name, *, kind='standard'):
        if name in self.components:
            # The real BookVBA's own wording, so a test can assert on it.
            raise ValueError(
                f"this workbook already has a VBA component named {name!r}")
        component = FakeVBComponentForForms(
            name, {'standard': 1, 'class': 2, 'form': 3}[kind])
        self.components[name] = component
        return component

    def remove_component(self, name):
        self.removed.append(name)
        self.components.pop(name, None)
        return self

    def project(self):
        return FakeVBProjectForForms(self.components)


class FakeApplicationForForms:
    """Every Run of the Form function hands back a NEW instance, so a test
    can tell "cached" from "fetched again"."""

    def __init__(self):
        self.runs = []
        self.instances = []

    def Run(self, macro):
        self.runs.append(macro)
        if 'WineOLE_Form_' not in macro:
            return None
        instance = FakeUserFormInstance()
        self.instances.append(instance)
        return instance


class FakeWorkbookForForms:
    def __init__(self):
        self.application = FakeApplicationForForms()

    def Application(self):
        return self.application

    def Name(self):
        return 'Book1'


class FakeBookForForms:
    """Stands in for the Book WRAPPER: `ole` and `vba` are what Forms and
    Form reach through, and both are plain attributes here (properties on the
    real one)."""

    def __init__(self):
        self.ole = FakeWorkbookForForms()
        self.vba = FakeBookVBAForForms()
