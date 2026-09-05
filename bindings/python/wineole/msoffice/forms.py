"""`book.forms`: the UserForms of a workbook.

A UserForm is a VBComponent of type 3 with a Designer; the form itself exists
only while VBA has it loaded, as its default instance. The wrapper reaches
that instance through four generated procedures in its own module, because a
UserForm's default instance has no COM name a client can ask for.
"""

from ..errors import RemoteError
from .controls import Controls, UserFormControls
from .passthrough import Passthrough


class Forms:
    """Reached as `book.forms`. Name note (measured against Excel 11):
    `forms` is free on Workbook."""

    USERFORM_TYPE = 3

    def __init__(self, book):
        self._book = book

    def add(self, name):
        """A new UserForm component, wrapped.

        The name is checked here as well as in add_component, because it also
        becomes part of four procedure names.
        """
        Controls.check_name(name)
        component = self._book.vba.add_component(name, kind='form')
        try:
            return Form(name, component, self._book)
        except Exception:
            # Form's helper write failed after the component was already
            # created; take the component back out rather than leaving it
            # behind with no helper to show or unload it.
            self._book.vba.remove_component(name)
            raise

    def __getitem__(self, name):
        """Re-bind an existing UserForm (one made earlier, or one in a
        workbook that was opened). None for a missing name or a component
        that is not a UserForm. VBAAccessDenied from `project` passes
        through.

        Only the LOOKUP is inside the try, deliberately: Form's constructor
        writes the helper block, and that write can be refused (a locked
        project, a refused AddFromString). Reporting a real failure as "not
        found" would send the caller looking for a form that is right there.
        """
        try:
            found = self._book.vba.project().VBComponents().Item(name)
            component = found if int(found.Type()) == self.USERFORM_TYPE else None
        except RemoteError:
            component = None

        if component is None:
            return None

        return Form(name, component, self._book)

    @classmethod
    def helper(cls, name):
        """The four generated procedures, with a trailing newline.

        `Show 0` is modeless, and it is the only form the wrapper offers: a
        modal Show blocks Excel's message loop, and with it the bridge --
        measured, the bridge freezes until the form closes.
        """
        return (f"Function WineOLE_Form_{name}() As Object\n"
                f"    Set WineOLE_Form_{name} = {name}\n"
                'End Function\n'
                f"Sub WineOLE_Show_{name}()\n"
                f"    {name}.Show 0\n"
                'End Sub\n'
                f"Sub WineOLE_Hide_{name}()\n"
                f"    {name}.Hide\n"
                'End Sub\n'
                f"Sub WineOLE_Unload_{name}()\n"
                f"    Unload {name}\n"
                'End Sub\n')


class Form(Passthrough):
    """One UserForm. `ole` is the Designer (the design-time form: Caption,
    Width, Height, and the Controls that `controls` wraps); `instance()` is
    the loaded form -- the object that shows, hides and fires events.

    Name note (measured): `name`, `component`, `ole`, `instance`, `show`,
    `hide`, `unload` are free on the Designer; `controls` is a deliberate
    shadow of `Designer.Controls`, which stays reachable as
    `form.ole.Controls()`.
    """

    def __init__(self, name, component, book):
        # `_ole` first, and underscored like every wrapper attribute:
        # Passthrough.__setattr__ sends anything else to COM.
        self._ole = component.Designer()
        self._name = name
        self._component = component
        self._book = book
        self._controls = None
        self._instance = None
        self._runtime = {}
        # On EVERY construction, add and re-bind alike. write is an upsert,
        # and this is what makes a form in a REOPENED workbook showable
        # without the caller remembering to do anything first.
        book.vba.write(Forms.helper(name), name=f"form_{name}")

    @property
    def name(self):
        return self._name

    @property
    def component(self):
        return self._component

    @property
    def controls(self):
        """The design-time controls on this form. Memoised, so `add` and the
        `[]` that finds the result share one collection."""
        if self._controls is None:
            self._controls = UserFormControls(self, self._book.vba)
        return self._controls

    def instance(self):
        """The default instance. Referencing it loads the form if it is not
        loaded (VBA auto-instantiation), so `shown()` before `show()` answers
        False and leaves the form loaded but hidden. Cached until
        `unload()`."""
        if self._instance is None:
            self._instance = self._run('Form')
        return self._instance

    def runtime_control(self, control_name):
        """The live counterpart of a design-time control, by name. Cached, so
        that `on` and the `off` that undoes it meet the same Events."""
        if control_name not in self._runtime:
            self._runtime[control_name] = self.instance().Controls().Item(
                control_name)
        return self._runtime[control_name]

    def show(self):
        """Modeless, always. Returns as soon as the form is on screen; the
        bridge stays responsive, which is what lets events reach Python."""
        self._run('Show')
        return self

    def hide(self):
        """Through VBA, like show and unload: the extender's Hide answers
        GetIDsOfNames and then refuses every out-of-process Invoke (measured,
        DISP_E_MEMBERNOTFOUND whatever the flags), while Visible reads
        fine."""
        self._run('Hide')
        return self

    def shown(self):
        return bool(self.instance().Visible())

    def unload(self):
        """Unloading destroys the runtime controls, and with them every event
        connection Python holds on them, so those are closed first -- on our
        side, deliberately, rather than left to fail when the object is gone.
        The next `instance()` loads a fresh form."""
        for control in self._runtime.values():
            control.ole_events.close()
        if self._instance is not None:
            self._instance.ole_events.close()
        self._runtime.clear()
        self._instance = None
        self._run('Unload')
        return self

    # --- private ----------------------------------------------------------

    def _run(self, verb):
        """Qualified with the workbook name, in single quotes, so the right
        book's procedure runs when several are open."""
        return self._book.ole.Application().Run(
            f"'{self._book.ole.Name()}'!WineOLE_{verb}_{self._name}")
