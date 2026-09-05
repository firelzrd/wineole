"""The Office wrapper: an addressing DSL, a correct bulk write, formatting
and Linux/Wine path conversion on top of the raw COM surface.

`import wineole` does NOT import this package -- it is a second, explicit
`from wineole.msoffice import Excel`, the same policy Ruby's separate
`require 'wineole/msoffice'` follows. The core is a general-purpose COM
bridge; someone who wants the bridge should not have to carry an Excel
wrapper to get it.
"""

from .passthrough import Passthrough
from .address import Address
from .paths import Paths
from .vba import VBA, VBAError, VBAAccessDenied
from .vba_block import VBABlock
from .vba_api import BookVBA, SheetVBA
from .controls import (ActiveXControls, Control, Controls, FormControls,
                       UserFormControls)
from .forms import Form, Forms
from .color import Color
from .format import Format
from .range import Range
from .sheet import Sheet
from .book import Book
from .excel import Excel

# The one entry point Format has, re-exported under a name that reads
# correctly at a call site holding a RAW COM range (from the passthrough,
# or straight from the client) rather than a wrapper Range.
apply_format = Format.apply

__all__ = [
    'Passthrough', 'Address', 'Paths', 'Color', 'Format', 'apply_format',
    'Range', 'Sheet', 'Book', 'Excel',
    'VBA', 'VBABlock', 'BookVBA', 'SheetVBA', 'VBAError', 'VBAAccessDenied',
    'Controls', 'Control', 'FormControls', 'ActiveXControls',
    'UserFormControls', 'Forms', 'Form',
]
