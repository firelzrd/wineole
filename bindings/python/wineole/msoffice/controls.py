"""The three families of clickable thing on an Excel sheet or UserForm.

`Controls` is the tables and the checks the three collections share; every
one of them either maps a Python-side name onto the COM name Excel wants, or
refuses a call before Excel is touched -- the wrapper's job is to make the
passthrough's traps unreachable, not to add features to Excel.

`Control` is one placed control from any family. It is the one wrapper in
this package that holds TWO COM handles, and the split is not cosmetic: a
worksheet ActiveX control is an OLEObject (Excel's host: Left, Top, Visible,
LinkedCell) around an MSForms control (Caption, Value, BackColor). `ole` is
the host, `ole_object` the MSForms control, and unknown names go to
`ole_object` because that is where Caption and Value live for every family.
"""

import re

from ..errors import RemoteError
from .passthrough import Passthrough
from .vba_api import BookVBA
from .vba_block import _chomp


class Controls:
    """A class used only through its classmethods; never instantiated -- the
    shape `VBA` and `VBABlock` already use."""

    # Legacy form controls: kind -> the Worksheet collection whose Add makes
    # one. Every collection takes (Left, Top, Width, Height). EditBox is
    # missing on purpose: it exists only on dialog sheets.
    FORM_KINDS = {
        'button': 'Buttons',
        'check_box': 'CheckBoxes',
        'option_button': 'OptionButtons',
        'list_box': 'ListBoxes',
        'drop_down': 'DropDowns',
        'spinner': 'Spinners',
        'scroll_bar': 'ScrollBars',
        'label': 'Labels',
        'group_box': 'GroupBoxes',
    }

    # Shape.FormControlType -> kind, for re-binding an existing shape. 3
    # (xlEditBox) is absent for the reason above, so it maps to None.
    FORM_CONTROL_TYPES = {
        0: 'button',
        1: 'check_box',
        2: 'drop_down',
        4: 'group_box',
        5: 'label',
        6: 'list_box',
        7: 'option_button',
        8: 'scroll_bar',
        9: 'spinner',
    }

    # MSForms 2.0: kind -> ProgID. Shorthand only -- any other string is
    # passed to Excel verbatim as a ProgID, which is how a control outside
    # this table gets placed.
    MSFORMS_KINDS = {
        'command_button': 'Forms.CommandButton.1',
        'text_box': 'Forms.TextBox.1',
        'combo_box': 'Forms.ComboBox.1',
        'list_box': 'Forms.ListBox.1',
        'check_box': 'Forms.CheckBox.1',
        'option_button': 'Forms.OptionButton.1',
        'toggle_button': 'Forms.ToggleButton.1',
        'spin_button': 'Forms.SpinButton.1',
        'scroll_bar': 'Forms.ScrollBar.1',
        'label': 'Forms.Label.1',
        'image': 'Forms.Image.1',
    }

    # A worksheet ActiveX control is two objects: the OLEObject Excel wraps
    # it in, and the MSForms control inside (OLEObject.Object). These keys
    # belong to the host; everything else goes inside.
    HOST_PROPS = ('linked_cell', 'list_fill_range', 'visible', 'print_object',
                  'placement')

    # A VBA identifier. Excel accepts any string as a shape name, but a name
    # that cannot appear in `Name_Click` is a control that can be placed and
    # never handled.
    VBA_NAME = re.compile(r'\A[A-Za-z][A-Za-z0-9_]{0,30}\Z')

    @classmethod
    def form_collection_for(cls, kind):
        """The legacy collection for a form-control kind.

        Ruby has a second branch here for a String kind ("form controls have
        no ProgID"); in Python a ProgID and a kind are both str, so there is
        nothing to tell apart. A ProgID gets the unknown-kind message, which
        already points at sheet.activex.
        """
        cls._require_str(kind, 'a control kind')
        collection = cls.FORM_KINDS.get(kind)
        if collection is not None:
            return collection

        raise ValueError(
            f"unknown form control kind {kind!r} -- expected one of "
            f"{', '.join(repr(key) for key in cls.FORM_KINDS)}. For an ActiveX "
            'control (events reach Python) use sheet.activex')

    @classmethod
    def progid_for(cls, kind):
        """A known kind's ProgID, or the string itself -- an unknown string
        IS a ProgID, and that escape hatch is the whole point. Ruby's
        "unknown ActiveX kind" message is unreachable in Python and is
        dropped."""
        cls._require_str(kind, 'a control kind')
        return cls.MSFORMS_KINDS.get(kind, kind)

    @classmethod
    def kind_for_progid(cls, progid):
        for kind, value in cls.MSFORMS_KINDS.items():
            if value == progid:
                return kind
        return progid

    @classmethod
    def check_name(cls, name):
        cls._require_str(name, 'a control name')
        if cls.VBA_NAME.fullmatch(name):
            return None

        raise ValueError(
            'name: must be a VBA identifier -- a letter, then letters, digits '
            'or underscores, at most 31 characters -- because it becomes the '
            f"`Name_Click` handler name. Got {name!r}")

    @classmethod
    def check_event(cls, event):
        cls._require_str(event, 'an event name')
        if cls.VBA_NAME.fullmatch(event):
            return None

        raise ValueError(
            "an event name is a VBA identifier such as 'Click' or 'KeyDown'. "
            f"Got {event!r}")

    @classmethod
    def check_free(cls, host, name, what):
        """Excel allows two shapes with the same name in silence, and a
        UserForm allows two controls with the same name in silence; after
        that `Name_Click` means either. A lookup that RAISES is the free
        case -- the return value of a successful one is never used."""
        try:
            host.Item(name)
        except RemoteError:
            return None

        raise ValueError(
            f"this {what} already has a control named {name!r}. Excel would "
            'add a second one silently and then Name_Click would be ambiguous; '
            'pick another name or remove the existing control first')

    @staticmethod
    def pascal(key):
        """snake_case -> PascalCase. COM matches names case-insensitively, so
        a key already in PascalCase survives the round trip well enough."""
        return ''.join(part.capitalize() for part in key.split('_'))

    @classmethod
    def put(cls, target, key, value):
        """One property assignment. Never wrapped: a refusal propagates as
        the RemoteError it is, and the caller's rollback deals with it."""
        setattr(target, cls.pascal(key), value)

    @classmethod
    def geometry(cls, *, sheet, at, left, top, width, height):
        """Either `at=` (a range on `sheet`, read for its box) or all four
        points. Never a mix, never a partial box, never a default --
        "wherever Excel puts it" stays a passthrough behaviour. `sheet=None`
        is a UserForm, which has no cells for `at=` to name.

        A refusal never touches Excel: every check that can fail runs before
        the one COM read in here.
        """
        points = {'left': left, 'top': top, 'width': width, 'height': height}
        given = [key for key, value in points.items() if value is not None]

        if at is not None and given:
            raise ValueError(
                'give either at= or left=/top=/width=/height=, not both (got '
                f"at={at!r} and {given!r})")

        if at is not None:
            if sheet is None:
                raise ValueError(
                    'a UserForm has no cells; give left=, top=, width= and '
                    'height= in points')

            box = sheet[at].ole
            return (box.Left(), box.Top(), box.Width(), box.Height())

        if not given:
            raise ValueError(
                "no position given: pass at='B2:C4' (a range on this sheet) or "
                'all four of left=, top=, width= and height= (points). There is '
                'no default position')

        if len(given) < 4:
            missing = [key for key, value in points.items() if value is None]
            raise ValueError(
                'left=, top=, width= and height= must all be given (missing '
                f"{missing!r})")

        return (left, top, width, height)

    @staticmethod
    def _require_str(value, what):
        if isinstance(value, str):
            return None
        raise TypeError(
            f"{what} must be a str, got {type(value).__name__}: {value!r}")


class Control(Passthrough):
    """One placed control, from any of the three families.

    TWO OBJECTS, ONE WRAPPER. `ole` is the host (a Shape, an OLEObject, or a
    design-time MSForms control) and `ole_object` the thing with Caption and
    Value; unknown names go to `ole_object` through the `_passthrough_target`
    override Passthrough reserves for exactly this. For the two non-ActiveX
    families they are the same object.

    Name note: `ole_object` carries the `ole_` prefix for the reason Proxy's
    meta-methods do -- a bare `object` would shadow COM `Object`, which on an
    MSForms control is a DIFFERENT thing (the raw control under the extender,
    with Caption but no Name and no events). `ctl.Object()` still reaches it.
    """

    def __init__(self, *, name, kind, family, ole, ole_object, vba, form=None):
        # `_ole` first: Passthrough.__setattr__ sends any name that does not
        # start with an underscore to COM, and every attribute here is
        # underscored for that reason.
        self._ole = ole
        self._ole_object = ole_object
        self._name = name
        self._kind = kind
        self._family = family
        self._writer = vba
        self._form = form

    @property
    def name(self):
        return self._name

    @property
    def kind(self):
        return self._kind

    @property
    def family(self):
        """'form_control', 'activex' or 'userform'."""
        return self._family

    @property
    def ole_object(self):
        """The thing with Caption and Value -- the inner MSForms control for
        an ActiveX host, the control itself otherwise."""
        return self._ole_object

    def _passthrough_target(self):
        return self._ole_object

    def events(self):
        """The Events object a callback would be registered on; None for a
        form control, which has none.

        Deliberately None rather than a raise, and deliberately different
        from `on`/`off`, which DO raise for that family: asking what an
        object has is not the same as trying to use what it does not have.
        """
        if self._family == 'activex':
            return self._ole_object.ole_events
        if self._family == 'userform':
            return self.runtime().ole_events
        return None

    def runtime(self):
        """A UserForm control's live counterpart on the form's default
        instance -- the object that fires events and shows changes while the
        form is loaded. The design-time control (`ole`) does neither."""
        if self._family != 'userform':
            raise ValueError(
                'only a UserForm control has a runtime instance; this is a '
                f"{self._family}")

        return self._form.runtime_control(self._name)

    def on(self, event, callback, *, args=True):
        return self._listenable().on(event, callback, args=args)

    def off(self, name_or_subscription):
        return self._listenable().off(name_or_subscription)

    def vba(self, event_or_body, body=None, *, params=None):
        """Write one VBA handler. A form control fires only Click, so it
        takes the body alone and is bound through OnAction; the other two
        take the event name, and Excel finds the procedure by its
        `Name_Event` name in the right module. `params=` is the parameter
        list, verbatim -- the wrapper carries no signature table.

        The block is named `Name_Event`, so writing the same event again
        replaces the handler (vba.write's own rule).
        """
        if self._family == 'form_control':
            if body is not None:
                raise ValueError(
                    'a form control fires only Click: call vba(body) with no '
                    'event name')

            macro = f"{self._name}_Click"
            self._writer.write(
                f"Sub {macro}()\n{self._indent(event_or_body)}\nEnd Sub",
                name=macro)
            self._ole.OnAction = macro
        else:
            if body is None:
                raise ValueError(
                    "vba(event, body) -- name the event, e.g. vba('Click', "
                    "'Range(\"A1\").Value = 1')")

            Controls.check_event(event_or_body)
            block = f"{self._name}_{event_or_body}"
            code = (f"Private Sub {block}({params or ''})\n"
                    f"{self._indent(body)}\nEnd Sub")
            if self._family == 'activex':
                self._writer.write(code, name=block)
            else:
                self._writer.write(code, name=block, into=self._form.name)
        return self

    # --- private ----------------------------------------------------------

    def _listenable(self):
        if self._family == 'form_control':
            raise ValueError(
                'form controls have no COM events; bind a macro with vba(...) '
                'or use sheet.activex for a control Python can listen to')

        return self.events()

    @staticmethod
    def _indent(body):
        """Four spaces on every non-empty line. `_chomp` drops ONE trailing
        terminator (the same rule vba_block writes with), and re.split keeps
        the trailing empty fields Ruby's `split(/\\r?\\n/, -1)` keeps, so a
        body ending in a blank line still does."""
        lines = re.split(r'\r?\n', _chomp(str(body)))
        return '\n'.join(line if line == '' else '    ' + line for line in lines)


class FormControls:
    """`sheet.form_controls`: the Forms-toolbar controls. Cheap to place and
    they save with the workbook, but they raise no COM events -- a handler is
    a macro named by OnAction, so `vba(body)` is the only way to react to
    one.

    `sheet` is the Sheet WRAPPER, not raw COM: `geometry` needs its `[]`.
    """

    def __init__(self, sheet):
        self._sheet = sheet
        self._writer = None

    def add(self, kind, *, name, at=None, left=None, top=None, width=None,
            height=None, **props):
        """Order matters and each step exists to make a specific passthrough
        trap unreachable: kind and name are checked before geometry so a typo
        fails without a round trip; the free-name check runs against Shapes,
        which sees every family; the rename and the properties run after Add
        with the new control deleted if either fails, so a refused property
        does not leave an unnamed button behind.

        `**props` are applied in call order (Python keeps it), each as one
        PascalCase COM property assignment.
        """
        collection = Controls.form_collection_for(kind)
        Controls.check_name(name)
        box = Controls.geometry(sheet=self._sheet, at=at, left=left, top=top,
                                width=width, height=height)
        Controls.check_free(self._sheet.ole.Shapes(), name, 'sheet')

        ole = getattr(self._sheet.ole, collection)().Add(*box)
        try:
            ole.Name = name
            for key, value in props.items():
                Controls.put(ole, key, value)
        except Exception:
            ole.Delete()
            raise
        return Control(name=name, kind=kind, family='form_control', ole=ole,
                       ole_object=ole, vba=self._book_vba())

    def __getitem__(self, name):
        """Re-bind a control already on the sheet. None when there is no
        shape of that name, or the shape is not a form control (an ActiveX
        shape raises on FormControlType; an EditBox maps to None).

        A miss is an answer, not an error, so this returns None rather than
        raising KeyError.
        """
        try:
            kind = Controls.FORM_CONTROL_TYPES.get(
                int(self._sheet.ole.Shapes().Item(name).FormControlType()))
            if kind is None:
                return None
            ole = getattr(self._sheet.ole, Controls.FORM_KINDS[kind])().Item(name)
        except RemoteError:
            return None

        return Control(name=name, kind=kind, family='form_control', ole=ole,
                       ole_object=ole, vba=self._book_vba())

    def _book_vba(self):
        """A form control's macro can live in any standard module, so it goes
        in the wrapper's own module of the parent workbook. Only `write` is
        used and paths never are, so convert_paths is moot -- and False keeps
        import_/export refusing on a wrapper nobody should reach them
        through."""
        if self._writer is None:
            self._writer = BookVBA(self._sheet.ole.Parent(), convert_paths=False)
        return self._writer


class ActiveXControls:
    """`sheet.activex`: OLEObjects hosting an MSForms control (or any other
    registered control, by ProgID). Each is two COM objects -- see Control --
    and the properties given at placement are routed accordingly: HOST_PROPS
    to the OLEObject, the rest inside.

    The five named arguments to OLEObjects.Add are not optional on Excel 11:
    with only Left and Top it fails with 0x800A03EC. Geometry has already
    guaranteed all four points by the time Add is called.
    """

    def __init__(self, sheet):
        self._sheet = sheet

    def add(self, kind, *, name, at=None, left=None, top=None, width=None,
            height=None, **props):
        progid = Controls.progid_for(kind)
        Controls.check_name(name)
        left_pt, top_pt, width_pt, height_pt = Controls.geometry(
            sheet=self._sheet, at=at, left=left, top=top, width=width,
            height=height)
        Controls.check_free(self._sheet.ole.Shapes(), name, 'sheet')

        ole = self._sheet.ole.OLEObjects().Add(
            ClassType=progid, Left=left_pt, Top=top_pt, Width=width_pt,
            Height=height_pt)
        try:
            ole.Name = name
            inner = ole.Object()
            for key, value in props.items():
                Controls.put(ole if key in Controls.HOST_PROPS else inner, key,
                             value)
        except Exception:
            ole.Delete()
            raise
        # The sheet's own VBA surface: Excel looks for a worksheet ActiveX
        # control's Name_Click in the sheet's module and nowhere else.
        return Control(name=name, kind=kind, family='activex', ole=ole,
                       ole_object=inner, vba=self._sheet.vba)

    def __getitem__(self, name):
        # Only the Item call is inside the try: progID() and Object() on an
        # object that was just found have no "not found" meaning to swallow.
        try:
            ole = self._sheet.ole.OLEObjects().Item(name)
        except RemoteError:
            return None

        return Control(name=name, kind=Controls.kind_for_progid(ole.progID()),
                       family='activex', ole=ole, ole_object=ole.Object(),
                       vba=self._sheet.vba)


class UserFormControls:
    """`form.controls`: MSForms controls on a UserForm, placed on the
    design-time Designer. Points only -- a form has no cells for `at=`.
    MSForms takes the name at Add time, so there is no rename step, and the
    box is four property puts after it (Controls.Add has no position
    arguments).
    """

    def __init__(self, form, book_vba):
        self._form = form
        self._book_vba = book_vba

    def add(self, kind, *, name, at=None, left=None, top=None, width=None,
            height=None, **props):
        progid = Controls.progid_for(kind)
        Controls.check_name(name)
        left_pt, top_pt, width_pt, height_pt = Controls.geometry(
            sheet=None, at=at, left=left, top=top, width=width, height=height)
        Controls.check_free(self._form.ole.Controls(), name, 'UserForm')

        ole = self._form.ole.Controls().Add(progid, name)
        try:
            ole.Left = left_pt
            ole.Top = top_pt
            ole.Width = width_pt
            ole.Height = height_pt
            for key, value in props.items():
                Controls.put(ole, key, value)
        except Exception:
            # By NAME, not by object: that is what a Designer's Controls
            # collection offers, unlike Shapes and OLEObjects, whose members
            # delete themselves.
            self._form.ole.Controls().Remove(name)
            raise
        return Control(name=name, kind=kind, family='userform', ole=ole,
                       ole_object=ole, vba=self._book_vba, form=self._form)

    def __getitem__(self, name):
        # kind is None: a placed MSForms control does not report the ProgID
        # it was made from.
        try:
            ole = self._form.ole.Controls().Item(name)
        except RemoteError:
            return None

        return Control(name=name, kind=None, family='userform', ole=ole,
                       ole_object=ole, vba=self._book_vba, form=self._form)
