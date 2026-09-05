"""The VBA surfaces of a workbook (`book.vba`) and of a worksheet
(`sheet.vba`).

TWO GRANULARITIES LIVE HERE and the names are what keep them apart.
`write` and `remove` act on a named BLOCK inside a module -- the span this
wrapper owns, delimited by sentinel comments, which may sit in a module full
of code somebody else wrote. `add_component`, `remove_component`, `import_`
and `export` act on whole COMPONENTS.

The verbs are not interchangeable either, and the difference is deliberate.
`write` is an upsert: writing the same name again replaces that block, so it
is idempotent. `add_component` is create-only and refuses a name that is
already taken -- "overwriting" a component would destroy whatever a person
put in it, so there is no way to ask for that.

WHERE CODE GOES DECIDES WHETHER IT CAN BE CALLED. Measured against a live
Excel 11:

  standard module   Run("Name") works and returns the value; a worksheet
                    formula =Name() works too. Private does NOT hide it
                    from either.
  ThisWorkbook,     Run("Name") fails ("macro not found"). Run with the
  a worksheet,      module qualified -- Run("Sheet1.Name") -- runs it but
  a UserForm        hands back None, so a Function's return value cannot be
                    collected. =Name() is #NAME?.

Two Ruby names are Python keywords, so they carry PEP 8's trailing
underscore: `import` is `import_` and `remove(name, from: x)` is
`remove(name, from_=x)`. Renaming them (`load`, `module=`) was the
alternative and was rejected so the README's Python twin reads line for
line against the Ruby original.
"""

import os
import tempfile

from ..errors import RemoteError
from .paths import Paths
from .vba import VBA, VBAError
from .vba_block import VBABlock


class BookVBA:
    """The VBA surface of a workbook, reached as `book.vba`."""

    # The module this wrapper makes for itself when `into=` is not given.
    DEFAULT_MODULE = 'WineOLE'

    # VBComponents.Add's type argument. 100 (a Document module -- a
    # worksheet or ThisWorkbook) is deliberately absent: Excel owns those
    # and neither creates nor destroys them on request.
    KINDS = {'standard': 1, 'class': 2, 'form': 3}

    # The Type of a component Excel owns. Cannot be added, cannot be removed
    # -- only emptied.
    DOCUMENT_TYPE = 100

    def __init__(self, ole, convert_paths):
        self._ole = ole
        self._convert_paths = convert_paths

    def write(self, code, *, name='main', into=None):
        """Put a named block of VBA into this workbook.

        `into=` names an EXISTING component -- a UserForm, ThisWorkbook, a
        worksheet's module, a module made with `add_component`. It is not
        created on demand: a typo would otherwise become a new module in
        silence. Without it the block goes in this wrapper's own module,
        which is created on demand because its name is not the caller's to
        get wrong.
        """
        VBABlock.write(self._target_module(into), name, code)
        return self

    def remove(self, name, *, from_=None):
        """Remove a named block.

        When `from_=` is given, only the block is removed -- a component the
        caller named is the caller's, never swept up here (a UserForm can be
        deleted, unlike ThisWorkbook, so this has to be a rule rather than
        an accident of what COM allows). Use `remove_component` to delete
        one deliberately.

        Without `from_=`, this wrapper's own module is the target, and when
        it has nothing but whitespace left the module goes too -- an empty
        module is litter, and that one IS ours to clean up.
        """
        if from_ is not None:
            VBABlock.remove(self._named_component(from_).CodeModule(), name)
            return self

        component = self._existing_component(self.DEFAULT_MODULE)
        if component is None:
            return self

        # `is not False`, not a truth test: an emptied module hands back [],
        # which is falsy in Python and truthy in Ruby. Treating [] as "there
        # was nothing to remove" would leave exactly the empty module this
        # branch exists to delete.
        remaining = VBABlock.remove(component.CodeModule(), name)
        if remaining is not False and VBABlock.blank_lines(remaining):
            self.project().VBComponents().Remove(component)
        return self

    def add_component(self, name, *, kind='standard'):
        """Create an empty component, and hand back the COM object.

        Create-only on purpose. The existence check has to come BEFORE Add,
        and that is measured rather than defensive: Add and the rename that
        follows it are not atomic. Adding under a taken name SUCCEEDS, and
        only the rename fails (0x80020009), leaving a stray `Module1` behind
        that nobody asked for. Checking first is what keeps the failure
        clean.
        """
        component_type = self.KINDS.get(kind)
        if component_type is None:
            raise ValueError(
                f"unknown component kind {kind!r} -- expected one of "
                f"{list(self.KINDS)!r}. A worksheet module and ThisWorkbook are "
                'not on that list because Excel owns them; they exist already '
                'and cannot be made')

        if self._existing_component(name) is not None:
            raise ValueError(
                f"this workbook already has a VBA component named {name!r}. "
                'add_component never overwrites one -- that would destroy '
                'whatever is in it. Remove it first, or use write(into=) to put '
                'a block inside it')

        component = self.project().VBComponents().Add(component_type)
        try:
            component.Name = name
        except RemoteError:
            # Add succeeded and the rename did not, so the component exists
            # under a name nobody chose. Take it back out rather than
            # leaving the litter the pre-check exists to prevent.
            self.project().VBComponents().Remove(component)
            raise ValueError(
                f"Excel refused {name!r} as a component name. VBA names start "
                'with a letter and hold letters, digits and underscores, up to '
                '31 characters') from None
        return component

    def remove_component(self, name):
        """Delete a component outright, with whatever is inside it."""
        component = self._named_component(name)
        if component.Type() == self.DOCUMENT_TYPE:
            raise ValueError(
                f"{name!r} is a module Excel owns (a worksheet's, or "
                "ThisWorkbook's) and cannot be deleted -- it exists for as long "
                "as the sheet or the workbook does. To take this wrapper's code "
                'back out of it, use remove(name, from_=)')

        self.project().VBComponents().Remove(component)
        return self

    def import_(self, path, *, encoding=None):
        """Read a VBA source file on this machine into the project as a new
        component. Excel's own Import does the work, so the component's name
        and kind come from the file.

        `encoding=` skips the detection when the caller already knows; the
        rules the default follows are on VBA.detect_encoding.
        """
        self._local_bridge('import')
        self._reject_dotdot(path)
        source = encoding or VBA.detect_encoding(path)
        with open(path, 'rb') as handle:
            raw = handle.read()
        text = self._decode(raw, source, path, guessed=encoding is None)

        with tempfile.TemporaryDirectory(prefix='wineole-vba') as directory:
            staged = os.path.join(directory, os.path.basename(path))
            with open(staged, 'wb') as handle:
                handle.write(self._to_codepage(text, path))
            self.project().VBComponents().Import(Paths.to_wine(staged))
        return self

    def export(self, name, path):
        """Write a component out as a file on this machine, in UTF-8 with LF
        -- what Excel produces is the ANSI codepage with CRLF, and the
        destination is a Linux path."""
        self._local_bridge('export')
        self._reject_dotdot(path)
        component = self._named_component(name)
        with tempfile.TemporaryDirectory(prefix='wineole-vba') as directory:
            staged = os.path.join(directory, os.path.basename(path))
            component.Export(Paths.to_wine(staged))
            with open(staged, 'rb') as handle:
                raw = handle.read()
            codepage = VBA.codepage()
            try:
                text = raw.decode(codepage)
            except UnicodeDecodeError as error:
                # Excel reported this codepage and then wrote a byte it does
                # not define. Substituting would hand back text that is
                # quietly wrong; a bare UnicodeDecodeError names neither the
                # component nor the codepage.
                raise VBAError(
                    f"could not decode the exported {name} as {codepage}: "
                    f"{error}") from None
            with open(path, 'w', encoding='utf-8', newline='') as handle:
                handle.write(text.replace('\r\n', '\n'))
        return self

    def project(self):
        """The workbook's VBA project, or an error that says what to do
        about it. The HRESULT and the message are both useless for telling
        this condition apart -- 0x800A03EC is what a rejected NumberFormat
        gives too, and the text is localized -- so the registry is what
        turns a refusal into advice.

        The try wraps ONE COM call. Anything wider would report an unrelated
        failure as "turn on AccessVBOM".
        """
        try:
            return self._ole.VBProject()
        except RemoteError:
            VBA.denied()

    # --- private ----------------------------------------------------------

    @staticmethod
    def _to_codepage(text, path):
        """The write direction, and it had been left bare: a file that reads
        cleanly can still hold a character the codepage cannot store. Same
        rule and same words as the string path -- refuse rather than let
        Excel substitute in silence."""
        try:
            return text.encode(VBA.codepage())
        except UnicodeEncodeError as error:
            VBA.unrepresentable(error.object[error.start], str(path))

    @staticmethod
    def _decode(raw, source, path, *, guessed):
        """Both encodings can be wrong at once: the bytes are not valid
        UTF-8 (which is what put us on the codepage branch) and not valid in
        the codepage either. Nothing can be inferred from that, so say so --
        a bare UnicodeDecodeError names neither the file nor why that
        encoding was the one tried."""
        try:
            text = raw.decode(source)
        except UnicodeDecodeError as error:
            if guessed:
                why = (
                    f"{source} was tried because the file has no BOM and its "
                    'bytes are not valid UTF-8, which rules UTF-8 out -- but they '
                    f"are not valid {source} either, so there is nothing left to "
                    'infer from. Pass `encoding:` if you know what this is')
            else:
                why = (f"you passed encoding: {source!r}, and the file's bytes "
                       'are not valid in it')
            raise ValueError(
                f"cannot read {path} as {source} ({error}). {why}") from None

        # A BOM the decode step turned into a character rather than
        # consuming (utf-8 keeps it; utf-8-sig would not).
        if text.startswith('\ufeff'):
            text = text[1:]
        return text

    def _local_bridge(self, what):
        """import_ and export hand Excel a path to a file on this machine.
        When the bridge is somewhere else that path means that machine's
        filesystem, and there is no sensible thing to do with it.

        Keying this off _convert_paths (rather than a separate "is this
        loopback" flag) is deliberate: it can only over-refuse a bridge that
        would have worked, never under-refuse one that would not.

        RuntimeError, not ValueError: it is the environment that is wrong,
        not the argument -- the same split Excel uses for "no active
        workbook".
        """
        if self._convert_paths:
            return
        raise RuntimeError(
            f"{what} needs the bridge to be on this machine: it stages a file "
            'and hands Excel the path, which means nothing on another host')

    @staticmethod
    def _reject_dotdot(path):
        """The basename of a path ending in ".." is literally "..". For
        import that path is read directly and raises IsADirectoryError of
        its own accord -- but a bare IsADirectoryError does not say why. For
        export the same basename feeds os.path.join(dir, "..") when staging,
        which resolves to the tmpdir's *parent*, and Export ends up trying
        to write a file over that directory. Nothing escapes and nothing is
        corrupted either way; this just gives both a clear error."""
        if os.path.basename(path) != '..':
            return
        raise ValueError(
            f"{path!r} is not a usable file path (its basename is \"..\")")

    def _existing_component(self, name):
        """None when the lookup refused. The try wraps ONE COM call: a
        project fetch that itself raised must surface as the denial it is,
        not as "that component is not there"."""
        components = self.project().VBComponents()
        try:
            return components.Item(name)
        except RemoteError:
            return None

    def _target_module(self, into):
        if into is None:
            return self._own_module().CodeModule()
        return self._named_component(into).CodeModule()

    def _own_module(self):
        """This wrapper's own module, made on demand. Unlike a name the
        caller passed, this one cannot be a typo, so creating it silently is
        safe."""
        found = self._existing_component(self.DEFAULT_MODULE)
        if found is not None:
            return found

        component = self.project().VBComponents().Add(self.KINDS['standard'])
        component.Name = self.DEFAULT_MODULE
        return component

    def _named_component(self, name):
        found = self._existing_component(name)
        if found is not None:
            return found

        raise ValueError(
            f"this workbook has no VBA component named {name!r}. Components are "
            'UserForms, ThisWorkbook, worksheet modules and standard modules; '
            'add one with add_component, or omit `into=` to use this wrapper\'s '
            'own module')


class SheetVBA:
    """The VBA surface of one worksheet, reached as `sheet.vba`.

    Blocks only. A worksheet's module cannot be created or deleted -- Excel
    makes it with the sheet and destroys it with the sheet -- so the
    component methods are absent here rather than present and always
    failing. Removing the last block empties the module; it does not remove
    it.
    """

    def __init__(self, ole):
        self._ole = ole

    def write(self, code, *, name='main'):
        VBABlock.write(self._code_module(), name, code)
        return self

    def remove(self, name):
        VBABlock.remove(self._code_module(), name)
        return self

    # --- private ----------------------------------------------------------

    def _code_module(self):
        """A worksheet's handlers live in the worksheet's own code module --
        that is where Excel looks for `<ActiveX control>_Click`. The module
        is named by the sheet's CodeName, inside the parent workbook's
        project.

        THE ORDER OF THESE TWO LINES IS LOAD-BEARING, and it is not obvious.
        Worksheet.CodeName comes back as "" until something has touched that
        workbook's VBProject -- measured: "" before, "Sheet3" after, and
        "Sheet3" on every read from then on. Reading the name first and then
        the project (which looks like the same code, and is what extracting
        a local naturally produces) hands VBComponents.Item("") and gets
        0x800A0009, "index out of range". So the project is fetched first,
        on purpose, and the name after it.

        The empty check is what keeps that from becoming a silent trap
        again: if this ever stops holding, it fails saying why instead of
        failing as a bare COM index error.
        """
        vb_project = self._project()
        code_name = self._ole.CodeName()
        if not code_name:
            raise VBAError(
                'this worksheet reports no CodeName even after its VBProject '
                'was opened, so there is no way to find its code module')

        return vb_project.VBComponents().Item(code_name).CodeModule()

    def _project(self):
        """Only the VBProject fetch is the denial. Wrapping the lookup that
        follows it in the same try would report any other COM failure -- a
        component that is not there, a module that will not open -- as "turn
        on AccessVBOM", which is advice for a condition the caller is not
        in."""
        try:
            return self._ole.Parent().VBProject()
        except RemoteError:
            VBA.denied()
