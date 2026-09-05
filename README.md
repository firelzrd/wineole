# WineOLE

Drive Windows OLE/COM automation (Excel, Word, ...) from Linux, through Wine.
A small Windows-side bridge process (`wineole-bridge`) runs under `wine` and
speaks a JSON Lines RPC protocol to native Python and Ruby clients, so host
code can script a Windows application as if it were a local object.

## Installation

Python:

```
pip install wineole
```

Ruby:

```
gem install wineole
```

Both packages bundle the prebuilt `wineole-bridge` binaries -- no separate
build step, and no Windows build of Python or Ruby is needed.

## Usage

The quickest way to get started -- open the default bridge and create a
COM object in one call each:

Python:

```python
import wineole

xl = wineole.create('Excel.Application')
xl.Visible = True
xl.Workbooks().Add()
xl.Worksheets()[1].Range('A1').Value = 'Hello from Python'
```

Ruby:

```ruby
require 'wineole'

xl = WineOLE.create('Excel.Application')
xl.Visible = true
xl.Workbooks.Add
xl.Worksheets[1].Range('A1').Value = 'Hello from Ruby'
```

`WineOLE.create`/`wineole.create()` (create a new instance), `.connect`/
`connect()` (attach to an already-running instance, failing if none exists),
and `.connect_or_create`/`connect_or_create()` (attach if something's
running, otherwise create -- the classic VBScript `GetObject`-or-
`CreateObject` idiom) all lazily open a default bridge connection the first
time any of them is called, then reuse it -- no separate client variable
needed for the common case. `WineOLE.close`/`wineole.close()` releases that
implicit default early if you want to (otherwise it lives for the process's
lifetime).

For anything beyond the zero-config case -- a specific host/port, an auth
token, more than one bridge connection at once -- use `Client.open`
explicitly and create objects on it directly:

Python:

```python
import wineole

client = wineole.Client.open()
xl = client.create('Excel.Application')

xl.Visible = True
xl.Workbooks().Add()
xl.Worksheets()[1].Range('A1').Value = 'Hello from Python'

client.close()
```

Ruby:

```ruby
require 'wineole'

client = WineOLE::Client.open
xl = client.create('Excel.Application')

xl.Visible = true
xl.Workbooks.Add
xl.Worksheets[1].Range('A1').Value = 'Hello from Ruby'

client.close
```

`WineOLE.open`/`wineole.open()` sits between the two: it's `Client.open`
reachable without typing `Client`, and it *also* becomes the implicit
default, so a `WineOLE.create`/`wineole.create()` call right after uses the
client you just explicitly configured rather than a second, separate
zero-config one.

The two clients are otherwise near-identical in capability -- the extra `()`
in the Python version is the one real syntactic difference: Python's
`__getattr__` can't tell a property read from a method call apart the way
Ruby's `method_missing` can, so every COM member access returns a callable
that has to actually be called. The snippets above are the whole raw
client surface; see `bindings/ruby/examples/` and
`bindings/python/examples/` for a larger, visual demo built on the Office
wrapper (below) instead.

### Meta-methods and the `ole_` prefix

A `Proxy` forwards any name it doesn't recognize straight to the remote COM
object, so its own bookkeeping methods are named with an `ole_` prefix --
`ole_handle`, `ole_session_id`, `ole_release`, `ole_leave_open`,
`ole_const_load`, and `ole_created?`/`ole_created` (below) -- to keep them
out of the way of a real COM member that happens to share the plain name.
`invoke` is the one deliberate exception: it's kept bare and public as an
explicit escape hatch for the rare case where a COM object really does
define a member called e.g. `ole_handle` --
`proxy.invoke('ole_handle', [], {})` reaches it
directly, bypassing the local method entirely. This mirrors real Ruby
`WIN32OLE`'s own `ole_*`-prefixed introspection methods and its own bare,
public `invoke`.

`create`/`connect`/`connect_or_create` set `ole_created?`/`ole_created` on
the `Proxy` they return: `true`/`True` if a new instance was created,
`false`/`False` if an existing one was attached to, and `nil`/`None` for
any object derived from another (`xl.Worksheets`, `xl.Workbooks.Add`, ...)
-- attach-vs-create isn't a meaningful question for those. It stays a plain
query. It used to be how you decided whether to `Quit` on the way out --
quit only what you created, never a human's already-open Excel:

Python:

```python
xl = wineole.connect_or_create('Excel.Application')
try:
    ...
finally:
    if xl.ole_created:
        xl.Quit()
```

Ruby:

```ruby
xl = WineOLE.connect_or_create('Excel.Application')
begin
  ...
ensure
  xl.Quit if xl.ole_created?
end
```

That manual pattern still works -- `Quit` is a direct COM call -- but it is
rarely the right tool now. Shutting an auto-created instance down is the
bridge's job, driven by the `cleanup:`/`cleanup=` steps declared when you
obtain it (see [Instance ownership and
cleanup](#instance-ownership-and-cleanup) below), and
`WineOLE::MSOffice::Excel`/`Excel.run` wire those up for you. Reach for
`ole_created?` when you genuinely need to branch on how you got the
instance, not as the standard shutdown hook.

`Client.open`/`wineole`'s equivalents reuse an already-running bridge on the
given port, or spawn one and wait for it to come up -- either language's
client can talk to a bridge the other one started. `client.close()`/
`client.close` (or the implicit default's own eventual cleanup) releases the
session; the bridge automatically frees everything the connection owned (COM
objects included) once it sees the connection drop, and shuts down after 30
minutes of no active connections.

### Instance ownership and cleanup

Quitting an auto-created Office instance is the bridge's job, not the
client's. The bridge identifies an instance across connections, so when
several programs share one it can tell which release is the last -- and only
that last release triggers shutdown. An instance is torn down exactly when
it was auto-created (someone `create`d it, or `connect_or_create` created it
rather than attaching) *and* its last user has released the root. A human's
Excel you reached with `connect` is never auto-created, so this never quits
it out from under them.

What "shutdown" runs is a list of **cleanup steps** you declare when you
obtain the instance, via `cleanup:` (Ruby) / `cleanup=` (Python) on
`Client#create`/`connect`/`connect_or_create` (and on `Proxy`'s equivalents
underneath). `steps` is a list of `[member_name, *args]`; a name ending in
`=` is a property set, everything else a method call. The bridge runs them
in order, best-effort -- a failing step is logged and the rest still run, so
`DisplayAlerts=` failing can't stop `Quit`. Arguments must be scalars
(bool/number/string/null); an object reference is rejected, since the
steps run after your connection may already be gone.

Python:

```python
client = wineole.open()
xl = client.create('Excel.Application', cleanup={
    'steps': [['DisplayAlerts=', False], ['Quit']],
})
```

Ruby:

```ruby
client = WineOLE.open
xl = client.create('Excel.Application', cleanup: {
  steps: [['DisplayAlerts=', false], ['Quit']],
})
```

Those exact steps are what `WineOLE::MSOffice::Excel` declares for you, so
its wrappers and `Excel.run` need no `cleanup:` of their own. (The
zero-config module-level `WineOLE.create`/`wineole.create` shortcuts take no
`cleanup:` -- go through a `Client` when you need to declare steps.)

**Leaving an instance open.** `ole_leave_open` on a `Proxy` (either
language), or `Excel#leave_open` on the Office wrapper, revokes the bridge's
permission to quit this instance: it keeps running after your program leaves
-- a finished report left on screen for a human, say. It is one-way; there
is no re-arm.

Python:

```python
xl.ole_leave_open()
```

Ruby:

```ruby
xl.ole_leave_open      # or, on the Office wrapper: excel.leave_open
```

**`on_cleanup`.** Alongside `steps`, a `cleanup:` (Ruby) / `cleanup=`
(Python) may carry an `on_cleanup` closure. When the last user leaves while
your client is still connected, the bridge asks the client to run this
closure FIRST, on the callback dispatcher thread, before running the steps.
It is a prelude, not a veto: after it returns, the steps STILL run, so the
closure cannot keep the instance alive by itself -- to do that, have it call
`ole_leave_open` (which revokes permission, so the steps are then skipped).
`ole_release` blocks until the closure and any steps that follow have
completed.

```ruby
xl = client.create('Excel.Application', cleanup: {
  steps: [['DisplayAlerts=', false], ['Quit']],
  on_cleanup: proc { xl.ole_leave_open },   # keep it alive instead of quitting
})
```

```python
xl = client.create('Excel.Application', cleanup={
    'steps': [['DisplayAlerts=', False], ['Quit']],
    'on_cleanup': lambda: xl.ole_leave_open(),   # keep it alive instead of quitting
})
```

**`InstanceClosingError`.** A `connect` that lands on an instance already in
the middle of shutting down raises `WineOLE::InstanceClosingError` (Ruby) /
`InstanceClosingError` (Python) rather than a generic `RemoteError` -- retry
to attach to a fresh one. `connect_or_create` never raises it: it detects
the closing instance and transparently creates a new one instead. One rare
consequence of that: a `connect_or_create` arriving in the brief window
while an `on_cleanup` closure is still deciding to `leave_open` may create a
second instance rather than attaching to the one being kept alive.

### Bulk range reads and writes

Reading or writing a multi-cell range goes over the wire as one array
rather than one call per cell, which matters a lot here: every property
access is a TCP round trip through Wine.

Python:

```python
sheet.Range('A1:C3').Value = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
rows = sheet.Range('A1:C3').Value()
```

Ruby:

```ruby
sheet.Range('A1:C3').Value = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
rows = sheet.Range('A1:C3').Value
```

Values inside the array are converted the same way scalars are -- a date
cell arrives as a `Time`/`datetime`, an empty cell as `nil`/`None`. Arrays
of three or more dimensions are not supported: writing one raises, and
reading one comes back as `{"$unsupported_vt": ...}` rather than being
silently flattened.

A `Time`/`datetime` can also be written -- as a scalar, or nested inside a
bulk array -- and it lands as a real Excel date, not text. The wall-clock
value goes across as-is: there is no timezone conversion (a COM date
carries no zone), and precision is to the second (no sub-second
component).

Python:

```python
sheet.Range('A1').Value = datetime.datetime(2026, 8, 31, 9, 30, 45)
sheet.Range('A1').Value()
# -> datetime.datetime(2026, 8, 31, 9, 30, 45)
```

Ruby:

```ruby
sheet.Range('A1').Value = Time.new(2026, 8, 31, 9, 30, 45)
sheet.Range('A1').Value
# -> 2026-08-31 09:30:45 +0900
```

The write sets the cell's *value*; what happens to its *display format*
depends on what the cell already had. Excel's own `Range.Value` setter
auto-detects a date the same way it does when a recognizable date is typed
into the UI: on a cell still at the default `General` format, writing a
date switches the format to a default date/time pattern on its own
(observed here as `"yyyy/m/d h:mm"`), so a freshly-written cell shows a
formatted date, not a serial number. If the cell already carries an
explicit non-date format from earlier (e.g. `'0.00'`), that format is left
untouched -- but then `Range.Value`'s own getter hands the value back
according to *that* format, i.e. a plain number, not a `Time`/`datetime`,
even though the same date is still stored underneath (`Range.Value2`
always returns the raw serial number, e.g. `46265.396...`, regardless of
format). A cell already formatted as **text** (`'@'`) is the case most
likely to surprise: the format is left untouched there too, but
`Range.Value`'s getter then hands back a *string* rendering of the date
(observed here as `"2026/8/31  9:30:45"`), not a `Time`/`datetime` and not
a number -- a column pre-formatted as text is the trap a script is most
likely to hit. This is Excel's own `Range.Value` contract, not something
wineole adds or controls -- set `NumberFormat` explicitly after writing if
a script needs the read-back type to stay a date no matter what format the
cell had before.

Excel's own `Range.Value` has a shape quirk worth knowing before you build
anything on top of it: a 1x1 range's `Value` is a bare scalar, not a
one-element array, no matter how the range was addressed --

```
Range("A1").Value    -> 42.0                 Range("A1:A1").Value -> 42.0     (scalar!)
Range("A1:B1").Value -> [[42.0, "right"]]
Range("A1:A2").Value -> [[42.0], ["below"]]
Range("J1:J1").Value -> nil                  (empty 1x1)
Range("J1:J2").Value -> [[nil], [nil]]
```

That's Excel's own behavior, not something wineole adds or could change --
any 1x1 range, however it was written, reads back as a scalar, and only a
range with more than one cell reads back as a list of rows.

## Office wrapper

Everything above is the raw COM surface: `Worksheets[1].Range('A1').Value`
and friends, exactly as `WIN32OLE` would spell it. The Office wrapper is a
small layer bundled with the same client that adds an addressing DSL, a
correct bulk-write, and Linux/Wine path conversion on top of that surface
-- for Excel specifically so far, though it is where a Word or PowerPoint
wrapper would land too. It is a second, separate import:

Ruby:

```ruby
require 'wineole/msoffice'
```

Python:

```python
from wineole.msoffice import Color, Excel
```

Deliberately **not** pulled in by `require 'wineole'` / `import wineole`
itself: the core is a general-purpose COM bridge, this part is
Office-specific and will only grow, and someone who wants the bridge
should not have to carry an Excel wrapper to get it.

The Ruby wrapper claims nothing in the top-level namespace -- everything
lives under `WineOLE::MSOffice`. Alias it yourself if you want the short
spelling:

```ruby
MSOffice = WineOLE::MSOffice
```

The two wrappers are the same wrapper in two idioms. Ruby's block forms
are Python context managers (`Excel.run`, `no_alert`, `no_update`), Ruby's
symbols are plain strings (`:center` is `'center'`), `range.to_a` is
`range.to_list()`, `book.each_sheet` is `book.sheets()`, `Color[...]` is
`Color.parse(...)`, and every COM member reached through the passthrough
needs the raw client's trailing `()` in Python.
Two VBA method names are Python keywords and carry PEP 8's
trailing underscore: `book.vba.import_(path)` and
`book.vba.remove(name, from_='Utils')`.

### Lifecycle: `Excel.run`

Ruby:

```ruby
WineOLE::MSOffice::Excel.run(:create) do |xl|
  # ...
end
```

Python:

```python
with Excel.run('create') as xl:
    ...
```

`run`'s first argument is `:create`, `:connect` or `:connect_or_create`
(default), matching the three ways to get hold of an `Excel.Application`
described above. On the way out its `ensure` calls `ole_release` -- it no
longer decides for itself whether to `Quit`. Each of `create`/`connect`/
`connect_or_create` declares Excel's cleanup steps (`DisplayAlerts = false`,
then `Quit`) to the bridge, and the bridge runs them when the last user of
an auto-created instance releases the root (see [Instance ownership and
cleanup](#instance-ownership-and-cleanup)). The observable outcome is
unchanged: `:create` always quits; `:connect` never quits a human's
already-open Excel; `:connect_or_create` quits exactly when it was the one
that created the instance. Attaching to somebody's already-open Excel and
then quitting it out from under them would throw away unsaved work, so this
is not a detail -- pinned in
`bindings/ruby/test/wineole/msoffice_integration_test.rb` against a real,
separately-started Excel instance, which the wrapper's `connect_or_create`
leaves running.

Two consequences of the bridge owning this rather than `run`: when several
programs share one instance, the LAST one to release it is the one that
quits it, not whoever created it; and `Excel#leave_open` revokes the
bridge's permission to quit, so a report meant to stay on screen for a human
keeps running after `run` returns.

(Wine's own COM teardown is not always prompt about it -- occasionally the
`EXCEL.EXE` process outlives `Quit` by well over 20 seconds before exiting
on its own, the same flake `wineole_integration_test.rb` documents for
connection teardown. The bridge still issues the `Quit` correctly; it is
Wine on the way out that is sometimes slow.)

Wrap any work that could pop a modal Excel dialog (saving, closing, format
warnings) in `xl.no_alert { ... }` -- under Wine a modal dialog is a hang,
not just an interruption, and every integration test in this suite wraps
its work in `no_alert` for exactly that reason. In Python it is a context
manager:

```python
with xl.no_alert():
    ...
```

### The addressing DSL

`[]`/`[]=` on `Excel` and `Sheet` both resolve through one address
grammar, `"[Book]Sheet!Range"`, with each part optional:

Ruby:

```ruby
xl['[:new]:new!A1'] = 'hello'   # new workbook + new worksheet + write A1
sheet = xl[':last!']            # -> Sheet: the worksheet just created
sheet['A1'].to_a                # -> [["hello"]] -- always 2-D
book = xl['[]']                 # -> Book: the active workbook
```

Python:

```python
xl['[:new]:new!A1'] = 'hello'   # new workbook + new worksheet + write A1
sheet = xl[':last!']            # -> Sheet: the worksheet just created
sheet['A1'].to_list()           # -> [['hello']] -- always 2-D
book = xl['[]']                 # -> Book: the active workbook
```

`Book` does not define `[]`/`[]=` -- from a `Book` reach a worksheet with
`book.sheet(name_or_index)`, which returns a `Sheet` you can index the same
way.

`:new` creates; `:first`/`:last` index by position; an empty part means
"the active one". A bare `sheet['A1']` (no `!`) addresses a cell on the
`Sheet` itself. Anything that isn't a recognized COM member on the wrapper
falls straight through to the underlying object -- `sheet.ole` for the raw
COM object explicitly, or just calling the member name (`sheet.PageSetup`)
since `method_missing` covers it either way.

### `write` vs. Excel's own assignment

This is the reason `Range#write` (and the `[]=` that uses it) exists.
Excel's own `Range.Value=` setter silently replicates a flat Array's first
element down every cell of a multi-cell range instead of laying the
values out -- no error, and nothing visible without opening the sheet.
Measured side by side, same shape, same source array, one call apart:

Ruby:

```ruby
sheet['A1:A10'] = (1..10).to_a
sheet['A1:A10'].to_a.flatten
# => [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]

sheet.ole.Range('B1:B10').Value = (1..10).to_a   # Excel's own assignment
sheet.ole.Range('B1:B10').Value.flatten
# => [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]
```

Python:

```python
sheet['A1:A10'] = list(range(1, 11))
[v for row in sheet['A1:A10'].to_list() for v in row]
# => [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0]

sheet.ole.Range('B1:B10').Value = list(range(1, 11))   # Excel's own assignment
# reads back as [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0]
```

`write` also refuses a value that doesn't fit the range's shape, rather
than truncating or leaving `#N/A` behind the way Excel's own setter does:

```ruby
sheet['E1:G3'] = [1, 2]
# ArgumentError: range is 3x3; a flat array only fits a single row or
# column -- pass rows, or use fill
```

```python
sheet['E1:G3'] = [1, 2]
# ValueError: range is 3x3; a flat list only fits a single row or
# column -- pass rows, or use fill
```

(`fill`, unlike `write`, is total -- it replicates/truncates/pads on
purpose, reproducing Excel's own column trap by design; reach for it only
when that's actually what's wanted.)

A single cell's `Value` stays Excel's own bare scalar (`sheet['A1'].Value
#=> 1.0`), which is exactly why `to_a` exists -- generic code that doesn't
know a range's size in advance wants `[[1.0]]`, not a branch of its own for
the 1x1 case.

### `Range#format`

`format` takes keyword arguments rather than a chain of verbs, because
formatting needs three states and a chain only has two: an **absent** key
leaves that attribute alone, and an explicit **`false`** turns it off. `nil`
means "I have no opinion" -- the same as omitting the key -- so
`format(bold: prefs[:bold])` is safe to write even when `prefs` has nothing
for `:bold`; it is not the same as passing `false`. Measured:

```ruby
sheet['A1'].format(bold: true)
sheet['A1'].format(bold: false)
sheet['A1'].format(bold: true)
sheet['A1'].format(bold: nil)      # leaves it alone -- not the same as false

prefs = {}
sheet['A2'].format(bold: prefs[:bold])
```

```
format(bold: true)   -> Font.Bold = true
format(bold: false)  -> Font.Bold = false
format(bold: nil)    -> Font.Bold = true   (left as it was)
format(bold: prefs[:bold]) with prefs empty -> no error; A2 (untouched) Font.Bold = false
```

The keys, and what they take:

| key | values |
| --- | --- |
| `bold:` | `true` / `false` |
| `italic:` | `true` / `false` |
| `underline:` | `true` / `false` / `:single` / `:double` / `:none` -- `true` is an alias for `:single` |
| `size:` | a number in `1..409`, Excel's own font size range |
| `color:` | `'#RRGGBB'`, `'#RGB'` or `[r, g, b]` / `false` to clear |
| `background:` | same values as `color:` / `false` to clear |
| `align:` | `:general` `:left` `:center` `:right` `:justify` |
| `valign:` | `:top` `:center` `:bottom` |
| `wrap:` | `true` / `false` |
| `number_format:` | a format code string, or `:general` / `:text` |
| `border:` | see below |

In Python the same keys take strings instead of symbols
(`underline='single'`, `align='center'`, `valign='top'`,
`border='outline'`, `border=['top', 'bottom']`,
`border={'edges': 'all', 'style': 'thick'}`), `False` still means "turn
it off" and `None` still means "no opinion". `number_format` takes a
format code string; `'general'` (any case) is the one that resolves
through `Application().International(26)`, and `'text'` is shorthand
for the `'@'` code, exactly as `:general` / `:text` are in Ruby.

`format` returns the `Range`, so it chains with `write`:

```ruby
sheet['A1:C1'].format(bold: true, align: :center).write(['Name', 'Qty', 'Price'])
```

```python
sheet['A1:C1'].format(bold=True, align='center').write(['Name', 'Qty', 'Price'])
```

`border:`'s value shape isn't guessable from the key name, so it gets its
own two forms. The simple form is a shorthand for the hash form, filled in
with `style: :thin`:

```ruby
sheet['A1:C3'].format(border: :all)      # every edge, including the inside ones
sheet['A1:C3'].format(border: :outline)  # the outer edges only
sheet['A1'].format(border: :bottom)      # one named edge, or an array of them
sheet['A1:C3'].format(border: false)     # clears every edge

# The hash form, for anything the shorthand can't say -- a style other
# than :thin, or a border colour:
sheet['A1:C3'].format(border: {edges: :all, style: :thick, color: '#999999'})
```

```python
sheet['A1:C3'].format(border={'edges': 'all', 'style': 'thick', 'color': '#999999'})
```

`edges:` takes `:all`, `:outline`, one of `:left` `:top` `:bottom` `:right`
`:inside_v` `:inside_h`, or an array of those (`[:top, :bottom]`). `style:`
takes `:none` `:thin` `:medium` `:thick` `:hairline` `:dash` `:dot`, and
defaults to `:thin` when the hash omits it. `color:` takes the same values
as the top-level `color:`/`background:` keys.

Colours are `'#RRGGBB'`, `'#RGB'` or `[r, g, b]`, converted through
`WineOLE::MSOffice::Color`, which is public and returns a plain `Integer` on
purpose. `format`'s own keys cover the colours it has names for, but COM's
colour surface cannot be enumerated -- the passthrough reaches every colour
property Excel has, including ones this wrapper will never grow a key for,
and they all take the same integer. A worksheet tab is one:

```ruby
WineOLE::MSOffice::Color['#FF0000']       #=> 255
WineOLE::MSOffice::Color[[255, 0, 0]]     #=> 255 -- same colour, the [r, g, b] form
WineOLE::MSOffice::Color.to_hex(255)      #=> '#FF0000' -- the inverse, for reading a colour back from COM
sheet.ole.Tab.Color = WineOLE::MSOffice::Color['#FF0000']
sheet.ole.Tab.ColorIndex                  #=> 3 -- Excel's own name for red
```

```python
Color.parse('#FF0000')       # => 255
Color.parse([255, 0, 0])     # => 255 -- same colour, the [r, g, b] form
Color.to_hex(255)            # => '#FF0000'
sheet.ole.Tab().Color = Color.parse('#FF0000')
```

`Interior.PatternColor` is another. That is the whole reason the conversion
is a public function returning an ordinary number rather than something
buried inside `format`: one function covers the entire surface, including
the parts nobody has thought of yet.

Excel stores colour as BGR, not RGB -- `0xFF0000` written raw to
`Interior.Color` reports back as blue, not red. That is exactly why
`Color[]` exists instead of a plain `.to_i(16)`, and why this suite's colour
test asserts on Excel's own `ColorIndex` rather than on the number that went
in: a reversed conversion round-trips through the number perfectly and would
only be caught by Excel actually naming the colour.

**The 56-colour palette.** The underlying file format is still Excel 2003's
(`.xls`), which has a fixed 56-colour palette and silently snaps anything
outside it to the nearest entry -- no error, no warning, and no way to tell
from the number handed to `Interior.Color` alone. Measured against a live
Excel 11:

| asked for | actually stored |
| --- | --- |
| `#FF0000` `#00FF00` `#0000FF` `#FFFFFF` `#000000` `#808080` | unchanged |
| `#EEEEEE` | `#FFFFFF` |
| `#123456` | `#003366` |
| `#E0FFC0` | `#CCFFCC` |

Not a wrapper defect -- Excel itself does the snapping; the wrapper only
hands the number through. `#EEEEEE` landing on pure white is the case worth
remembering: a "light grey" that was meant to be subtle turns invisible.

**`number_format: :general`.** The obvious spelling fails outright on this
(Japanese-localized) Excel:

```ruby
sheet['C2'].ole.NumberFormat = 'General'
# WineOLE::RemoteError: WIN32OLERuntimeError: <the Range class's NumberFormat
# property cannot be set, in this Excel's own language> (0x800A03EC)
```

`'General'`, the English word, is not a locale-neutral format code, and the
the translated spelling that does work on any one machine is not portable
to a different locale either. `number_format: :general` sidesteps both by asking
Excel itself, through `Application.International(26)`, which returns
whichever spelling this particular Excel wants -- and resets a cell to
exactly the format a pristine cell already has:

```ruby
sheet['C1'].format(number_format: '0.00')     # NumberFormat => "0.00"
sheet['C1'].format(number_format: :general)   # NumberFormat => whatever this Excel calls it
```

```python
sheet['C1'].format(number_format='0.00')      # NumberFormat => "0.00"
sheet['C1'].format(number_format='general')   # NumberFormat => whatever this Excel calls it
```

Writing a date does not get a guessed format from this wrapper. Excel
itself picks something at write time based on locale and cell history --
measured here, a bare `Time` landed on `"yyyy/m/d"` -- and that default is
not something to build on; a different cell, locale or Excel version can
pick differently. Ask for the format explicitly instead of relying on
whatever Excel guessed:

```ruby
sheet['D1'].write(Time.new(2026, 1, 15))
sheet['D1'].format(number_format: 'yyyy/m/d')
sheet['D1'].ole.Text   #=> "2026/1/15"
```

### Paths: `save_as` and `local_path`

`Book#save_as` and `Book#local_path` take/return **Linux** paths, converted
through `winepath` under the hood (only when the bridge is on the same
loopback machine as the client -- a remote bridge's filesystem means
nothing here, so paths pass through unconverted in that case):

Ruby:

```ruby
book.save_as('/tmp/wineole-readme.../out.xls')
File.exist?('/tmp/wineole-readme.../out.xls')  #=> true
book.local_path                                #=> "/tmp/wineole-readme..."
```

Python:

```python
book.save_as('/tmp/wineole-readme.../out.xls')
book.local_path              # => "/tmp/wineole-readme..."
```

`local_path` mirrors COM's own `Workbook.Path` -- the containing folder,
not the file itself; get the file's own path with `book.local_file`, gated
the same way `local_path` is (calling
`WineOLE::MSOffice::Paths.to_local(book.FullName)` directly skips that gate
and, over a remote bridge, silently converts a path that was never local
to begin with).

### Running VBA

Most of what people reach VBA for does not need VBA at all. These work on a
stock Excel with nothing switched on:

```ruby
xl.ole.Evaluate('1+2')                 #=> 3.0
xl.ole.WorksheetFunction.Sum(1, 2, 3)  #=> 6.0
xl.ole.ExecuteExcel4Macro('1+2')       #=> 3.0
```

```python
xl.ole.Evaluate('1+2')                   # 3.0
xl.ole.WorksheetFunction().Sum(1, 2, 3)  # 6.0
xl.ole.ExecuteExcel4Macro('1+2')         # 3.0
```

Reach for VBA when you need something the object model cannot express, or
when a loop is too slow across the bridge -- code running inside Excel pays
no round trip per cell.

**Project access is off by default.** Writing to a workbook's VBA project
needs `AccessVBOM` in the registry, which Office ships disabled. The gem and the Python
package both install `wineole-vba` to read and set it -- one command, the
same output either way:

```
$ wineole-vba
VBA project access: enabled

$ wineole-vba enable
VBA project access: enabled
Restart Excel for this to take effect -- it reads the setting at startup.
```

That restart line is not boilerplate: Excel reads the setting once, when it
starts, so flipping it under a running Excel changes nothing until that
Excel exits. The same three states are readable from Ruby --
`WineOLE::MSOffice::VBA.state` returns `:enabled`, `:disabled` or `:unset`
(the value absent entirely), with `enabled?`, `enable!` and `disable!`
alongside it. `enable!` and `disable!` return false rather than raising if
`wine` is not on `PATH`. Python spells the same three the same way, minus
Ruby's bangs and question mark: `VBA.state()` returns `'enabled'`,
`'disabled'` or `'unset'`, with `VBA.enabled()`, `VBA.enable()` and
`VBA.disable()` alongside it, from `from wineole.msoffice import VBA`.

This is a macro security setting, and it is machine-wide: turning it on
lets *any* Office automation on this machine reach VBA projects, not only
this library. It is yours to turn back off.

#### Blocks and components

`book.vba` is the whole VBA surface, and it works at two levels.

A **block** is a named span of code inside a module, delimited by sentinel
comments. It is what this wrapper owns; the module around it may hold code
somebody wrote by hand, which is never touched.

```ruby
book.vba.write(<<~VB, name: 'helpers')
  Function Doubled(a)
    Doubled = a * 2
  End Function
VB

xl.ole.Run('Doubled', 21)   #=> 42
```

```python
book.vba.write("""
Function Doubled(a)
  Doubled = a * 2
End Function
""", name='helpers')

xl.ole.Run('Doubled', 21)   # 42
```

Writing the same name again replaces just that block. Names match
case-insensitively, as VBA's own identifiers do, so `helpers` and `Helpers`
are one block and not two colliding ones. `remove` takes it back out, and
deletes this wrapper's own module once nothing is left in it -- removing
one function does not cost you the module, and an emptied module does not
linger:

```ruby
book.vba.write(one, name: 'helpers')
book.vba.write(two, name: 'formatting')
book.vba.remove('helpers')       # 'formatting' still there, module still there
book.vba.remove('formatting')    # module gone
```

```python
book.vba.write(one, name='helpers')
book.vba.write(two, name='formatting')
book.vba.remove('helpers')       # 'formatting' still there, module still there
book.vba.remove('formatting')    # module gone
```

A **component** is a whole module, class module or UserForm:

```ruby
book.vba.add_component('Utils')                       # kind: :standard by default
book.vba.write(code, name: 'x', into: 'Utils')
book.vba.remove_component('Utils')                    # with everything in it
```

```python
book.vba.add_component('Utils')                       # kind='standard' by default
book.vba.write(code, name='x', into='Utils')
book.vba.remove_component('Utils')                    # with everything in it
```

The verbs are not interchangeable, and the difference is the point.
`write` is an upsert -- writing the same block name again replaces it, so
it is idempotent. `add_component` is create-only and refuses a name that is
taken, because "overwriting" a component would destroy whatever a person
put in it. `into:` will not create a module either: a typo would otherwise
become a new module in silence.

`kind:` takes `:standard`, `:class` or `:form` (`kind='standard'`,
`'class'`, `'form'` in Python). A worksheet's module and
`ThisWorkbook` are not on that list -- Excel makes those with the sheet and
the workbook, and `remove_component` refuses them for the same reason.
Removing the last block from one empties it instead.

Code injected this way survives a save and reopen, as does a UserForm you
add to the project.

#### Where code goes decides whether it can be called

This is not a style preference. Measured against a live Excel 11:

| where the code is | `Run("Name")` | `Run("Module.Name")` | `=Name()` in a cell |
|---|---|---|---|
| standard module | value comes back | -- | value comes back |
| `ThisWorkbook` | macro not found | runs, **returns nil** | `#NAME?` |
| a worksheet's module | macro not found | runs, **returns nil** | `#NAME?` |
| a UserForm | macro not found | runs, **returns nil** | `#NAME?` |

`Private` does **not** hide a procedure in a standard module: it answers a
bare `Run` and a worksheet formula just the same as `Public` does.

So code you mean to call belongs in a standard module, which is where
`into: nil` puts it. Code in a sheet or `ThisWorkbook` module is there for
Excel to reach -- an ActiveX control's `_Click` (see [Controls](#controls)),
a workbook event -- and a
`Function` there cannot hand a value back through `Run` at all.

Handlers go in the sheet's own module:

```ruby
sheet.vba.write(code, name: 'guard')
sheet.vba.remove('guard')
```

```python
sheet.vba.write(code, name='guard')
sheet.vba.remove('guard')
```

`sheet.vba` has no component methods rather than having ones that always
fail: a worksheet's module cannot be created or deleted.

Name procedures with something that cannot be read as a cell address, and
not after their own module. A function called `F1` or `C3` fails to run
because Excel resolves the name as a reference first; a function with the
same name as its module loses to the module.

#### Files in and out

`import` and `export` take Linux paths. `export` writes UTF-8 with LF:

```ruby
book.vba.export('WineOLE', '/tmp/WineOLE.bas')
book.vba.import('/tmp/WineOLE.bas')
book.vba.import(path, encoding: 'CP932')   # skip detection when you know
```

```python
book.vba.export('WineOLE', '/tmp/WineOLE.bas')
book.vba.import_('/tmp/WineOLE.bas')
book.vba.import_(path, encoding='cp932')   # skip detection when you know
```

`import` is a Python keyword, so the method is `import_`; `remove`'s
`from:` is `from_=` for the same reason. Python's encoding names are
codec names -- `'cp932'`, not `'CP932'` -- and `VBA.codepage()` returns
one.

`import` decides the encoding on evidence, never on a heuristic. A BOM is
conclusive and wins. Otherwise, bytes that are not valid UTF-8 *prove* the
file is not UTF-8, so it is read as the machine's ANSI codepage -- which is
what Excel's own Export writes and what a `.bas` from any Windows toolchain
is. Anything else is read as UTF-8.

The direction matters and is why the codepage is never the default:
measured on a CP932 host, a codepage file read as UTF-8 comes out invalid
95% of the time at a single non-ASCII character and 99.99% by five, so real
codepage files are caught almost without exception -- while UTF-8 bytes
read as the codepage come out *valid but wrong* from two characters on.
Guessing UTF-8 and being wrong is loud; guessing the codepage and being
wrong is silent. What is left over -- a codepage file whose bytes happen to
be valid UTF-8 -- is undecidable for anyone, and `encoding:` is there for
it.

#### The codepage limit

Excel stores a module's text in the machine's ANSI codepage, not Unicode.
Whether a character survives therefore depends on the machine, and on
nothing about this library: on a CP932 host, Japanese and full-width text
round-trips exactly, while `é`, `✓` and simplified Chinese have no
representation at all and Excel replaces them silently -- `café` comes back
`cafe`, `✓` comes back `?`.

Silent replacement is worse than an error, so this wrapper refuses instead:

```ruby
book.vba.write('x = "café"', name: 'bad')
# ArgumentError: this code contains "é", which the system codepage (CP932)
# cannot represent. Excel stores a module's text in that codepage, so the
# character would be silently replaced rather than stored. Rewrite it with
# Chr()/ChrW() escapes, which are built at run time and are not bound by
# the codepage the source text is
```

```python
book.vba.write('x = "café"', name='bad')
# ValueError: this code contains 'é', which the system codepage (cp932)
# cannot represent. Excel stores a module's text in that codepage, so the
# character would be silently replaced rather than stored. Rewrite it with
# Chr()/ChrW() escapes, which are built at run time and are not bound by
# the codepage the source text is
```

`import` refuses the same thing for the same reason, and says so in the
same words -- only the subject changes, from "this code" to the file's
path. Both paths hand text to Excel and both are bound by the one
codepage, so which one you hit is not something the message should make
you work out.

`ChrW()` is the way through: the *source text* is limited to the codepage,
but a string a running macro builds is not. `ChrW(233)` produces `é` at
runtime, in a module whose text stayed pure ASCII.
`WineOLE::MSOffice::VBA.codepage` -- `VBA.codepage()` in Python -- reports
what this machine can hold.

### Controls

Excel has three kinds of clickable thing, and they differ in ways that
matter before a single one is placed:

| | `sheet.form_controls` | `sheet.activex` | `book.forms['F'].controls` |
|---|---|---|---|
| what it is | Forms-toolbar control (`Buttons`, `CheckBoxes`, ...) | an MSForms control hosted in an `OLEObject` | an MSForms control on a UserForm |
| lives on | a worksheet | a worksheet | a UserForm (`VBComponent` type 3) |
| COM events reach the client | **no** -- a macro named by `OnAction` | yes, from `OLEObject.Object` | yes, from the loaded form's control |
| VBA handler lives in | any standard module (`Name_Click`, bound via `OnAction`) | the **sheet's** module (`Private Sub Name_Click`) | the **form's** module |
| placement cost (Excel 11 under Wine) | ~3 ms (VBA spec) / 5 ms here | ~31 ms (VBA spec) / 364 ms here | ~10 ms (VBA spec) / 11 ms here |
| saved with the workbook | yes | yes | yes (`.xlsm`) |

"Here" is one `add` through the wrapper and the bridge plus the reads that
check it (name, box, caption), averaged over every kind of the family in
one run of the integration test; the VBA numbers are Excel alone.

Each is placed through the same call shape, `add(kind, name:, position,
**properties)`:

```ruby
sheet = xl[':first!']

go = sheet.form_controls.add(:button, name: 'Go', at: 'B2:C3', caption: 'Go')
ok = sheet.activex.add(:command_button, name: 'OK', left: 10, top: 10, width: 90, height: 24,
                       caption: 'OK', back_color: 0x00FF00)

form = xl['[]'].forms.add('AppForm')
name = form.controls.add(:text_box, name: 'Name', left: 12, top: 12, width: 160, height: 20)
```

```python
sheet = xl[':first!']

go = sheet.form_controls.add('button', name='Go', at='B2:C3', caption='Go')
ok = sheet.activex.add('command_button', name='OK', left=10, top=10, width=90, height=24,
                       caption='OK', back_color=0x00FF00)

form = xl['[]'].forms.add('AppForm')
name = form.controls.add('text_box', name='Name', left=12, top=12, width=160, height=20)
```

**Position** is either `at:` (a range on that sheet; the control takes the
range's `Left`/`Top`/`Width`/`Height`) or all four of `left:`, `top:`,
`width:`, `height:` in points. Not both, not three of the four, and there
is no default: a control with no position is an `ArgumentError` (a
`ValueError` in Python) before Excel is touched. UserForms have no cells,
so their controls take points only.

**Kinds.** Form controls: `:button :check_box :option_button :list_box
:drop_down :spinner :scroll_bar :label :group_box`. ActiveX and UserForm
controls (MSForms 2.0): `:command_button :text_box :combo_box :list_box
:check_box :option_button :toggle_button :spin_button :scroll_bar :label
:image`. Anything else registered on the host is reached by its ProgID as a
String, `sheet.activex.add('MSComctlLib.ProgCtrl.2', name: 'Ext', at: 'D2')`
-- this host has the Windows Common Controls registered; what yours has is
environment-specific, and `kind` reports the String back.

**Names** become VBA identifiers (`Name_Click`), so `name:` must be one: a
letter, then letters, digits or underscores, at most 31 characters. Excel
would let two shapes share a name in silence; the wrapper refuses the
second one across all families on a sheet, and across a form's controls.

**Properties** given at placement are put after the control is named, as
`PascalCase` COM property assignments (`back_color: 255` is
`.BackColor = 255`). On a worksheet ActiveX control `linked_cell`,
`list_fill_range`, `visible`, `print_object` and `placement` go to the
`OLEObject` host; everything else goes to the MSForms control inside it.
A property Excel rejects raises the original `WineOLE::RemoteError` and
the control that was just added is deleted again, so a typo does not leave
an unnamed control behind.

The result is a `Control`: `name`, `kind`, `family` (`:form_control`,
`:activex`, `:userform`), `ole` (the COM object Excel handed back -- the
`OLEObject` host for ActiveX) and `ole_object` (the thing with `Caption`
and `Value`; the same as `ole` except for ActiveX). Unknown methods go to
`ole_object`, so `ok.Caption = 'Run'` and `ok.Value` do what they read as;
host members are reached through `ok.ole` (`ok.ole.LinkedCell = 'A1'`).
`ole_object` is prefixed like every meta-method here (see [Meta-methods
and the `ole_` prefix](#meta-methods-and-the-ole_-prefix)): an MSForms
control has a COM member called `Object` of its own -- the raw control
under the extender, with a `Caption` but no `Name` and no events -- and
`ok.Object` still reaches it.
Re-bind an existing control with `[]`: `sheet.activex['OK']`,
`sheet.form_controls['Go']`, `form.controls['Name']` -- `nil` (`None` in
Python) when there is none of that name in that family.

#### Handlers: a callback or VBA

```ruby
ok.on('Click') { puts 'clicked' }          # runs in this process
ok.vba('Click', 'Range("A1").Value = 1')   # runs inside Excel
go.vba('Range("A2").Value = 2')            # form control: Click only, via OnAction
```

```python
ok.on('Click', lambda *args: print('clicked'))   # runs in this process
ok.vba('Click', 'Range("A1").Value = 1')         # runs inside Excel
go.vba('Range("A2").Value = 2')                  # form control: Click only, via OnAction
```

`on(event, args: true) { |*args| }` is the [COM events](#com-events)
machinery on the right object -- `OLEObject.Object` for a worksheet ActiveX
control, the loaded form's control for a UserForm -- and `off` undoes it --
in Python the block is a positional `callback`: `on(event, callback, args=True)`.
Form controls have no COM events, so `on` raises and says to use `vba` or
`sheet.activex`.

`vba(event, body, params: nil)` writes `Private Sub Name_Event(params)`
into the module Excel looks in (the sheet's, or the form's) as a named
block, so writing the same event again replaces the handler. `params:` is
the parameter list verbatim, e.g. `'ByVal KeyCode As
MSForms.ReturnInteger, ByVal Shift As Integer'` for `KeyDown`. A form
control fires only `Click`, so it takes the body alone: `vba(body)` writes
`Sub Name_Click()` to the wrapper's own standard module and sets
`OnAction`. Bodies go through the same codepage check as `vba.write`.

#### UserForms

```ruby
form = xl['[]'].forms.add('AppForm')          # a new UserForm component
ok = form.controls.add(:command_button, name: 'OK', left: 10, top: 40, width: 80, height: 24)
ok.on('Click') { |*| form.hide }

form.show                                    # modeless; returns at once
form.shown?                                  #=> true
ok.runtime.Caption = 'Go'                    # the live control, while loaded
form.hide
form.unload                                  # closes the subscriptions, unloads
book.forms['AppForm']                        # re-bind later; nil if none
```

```python
form = xl['[]'].forms.add('AppForm')             # a new UserForm component
ok = form.controls.add('command_button', name='OK', left=10, top=40, width=80, height=24)
ok.on('Click', lambda *args: form.hide())

form.show()                                      # modeless; returns at once
form.shown()                                     #=> True
ok.runtime().Caption = 'Go'                      # the live control, while loaded
form.hide()
form.unload()                                    # closes the subscriptions, unloads
book.forms['AppForm']                            # re-bind later; None if none
```

A UserForm exists twice: as its **Designer** (`form.ole`; what `controls`
edits, what is saved) and as its **default instance** (`form.instance`;
what is shown and what fires events). The wrapper reaches the instance
through four procedures it writes into its own `WineOLE` module,
`WineOLE_Form_<Name>`, `WineOLE_Show_<Name>`, `WineOLE_Hide_<Name>` and
`WineOLE_Unload_<Name>`,
because a form's instance has no COM name a client can ask for.
`ctl.runtime` is a control's live counterpart on that instance; `on`
registers there.

#### Constraints

- `OLEObjects.Add` with `Left`/`Top` but no `Width`/`Height` fails with
  `0x800A03EC` on Excel 11; pass all four or none. The wrapper always
  passes four.
- UserForms: `Show 0` only. A modal `Show` blocks Excel's message loop,
  and with it the bridge, until the form closes.
- `MsgBox`, `InputBox` and any modal dialog inside a VBA handler freeze the
  bridge for the same reason. Write to a cell, `Debug.Print`, or handle it
  in Ruby.
- Save as `.xlsm` to keep handlers and the generated `WineOLE_*` block;
  `.xlsx` drops them silently under `DisplayAlerts = False`.
- The generated `form_<Name>` block ships with the workbook; deleting it is
  safe, `forms['<Name>']` regenerates it.
- UserForm design-time property changes appear only after `unload`; use
  `ctl.runtime` while the form is loaded.
- A String ProgID works if that control is registered on the host; what
  Wine has registered is environment-specific.
- The wrapper never toggles Excel's design mode; `OLEObjects.Add` from
  automation does not enter it (measured).

## COM events

A COM object that raises events -- an `Application`, a `Workbook`, a
`Worksheet`, a UserForm control -- hands them to callbacks you register on
its `ole_events`:

```ruby
xl = WineOLE.create('Excel.Application')
xl.Visible = false
xl.DisplayAlerts = false
book = xl.Workbooks.Add
sheet = book.Worksheets(1)

xl.ole_events.on('SheetChange') do |sh, target|
  puts "#{sh.Name}!#{target.Address} is now #{target.Value.inspect}"
end

sheet.Range('A1').Value = 42
```

```
Sheet1!$A$1 is now 42.0
```

```python
xl = wineole.create('Excel.Application')
xl.Visible = False
xl.DisplayAlerts = False
book = xl.Workbooks().Add()
sheet = book.Worksheets()[1]

def on_change(sh, target):
    print(f"{sh.Name()}!{target.Address()} is now {target.Value()!r}")

xl.ole_events.on('SheetChange', on_change)

sheet.Range('A1').Value = 42
```

```
Sheet1!$A$1 is now 42.0
```

Callbacks arrive on another thread, so every example below assumes the
process is still alive when the event lands -- an `irb` session, a program
with something else to do, or a `sleep`/`Queue#pop` after the write. A script
that exits the instant it writes the cell can exit before the callback runs.

`ole_events` carries the `ole_` prefix for the reason every meta-method here
does (see [Meta-methods and the `ole_` prefix](#meta-methods-and-the-ole_-prefix)) --
a `Proxy` forwards unknown names straight to COM, so a bare `events` would
shadow a real `Events` member on some object that has one.

**Registering a callback is all there is.** There is no `subscribe` step, no
`Advise` to arrange, and nothing you have to shut down afterwards. The first
callback for an event subscribes on the bridge, which advises the COM
source; removing the last callback *for that event* unsubscribes, and
removing the last callback *on that object* takes the `Advise` and the
object's place on the connection with it. The dispatcher thread belongs to
the connection rather than to any one object, so it goes with the last
callback *on the connection*. Splitting those apart would allow the state
where a callback is registered and its event never arrives, with nothing to
show for it.

**A callback may call COM freely.** The example above already does -- both
`sh.Name` and `target.Value` are round trips to Excel made from inside the
callback. This works because the callback does not run on the thread that
reads the connection: the client's reader thread hands each event frame to a
dispatcher thread and goes straight back to reading, so the response to the
callback's own call still has somebody waiting for it. (Run callbacks on the
reader thread instead and the first COM call from inside one wedges the whole
connection -- the callback waits for an answer nobody is left to read.)

**Callbacks run on one dispatcher thread per connection, in arrival order,
one at a time.** Two callbacks on one event, and two changes:

```ruby
first  = xl.ole_events.on('SheetChange') { |_sh, t| puts "first saw #{t.Address}" }
second = xl.ole_events.on('SheetChange') { |_sh, t| puts "second saw #{t.Address}" }
sheet.Range('B1').Value = 1
sheet.Range('B2').Value = 2
```

```
first saw $B$1
second saw $B$1
first saw $B$2
second saw $B$2
```

One thread means a slow callback delays every event behind it, and that a
callback never runs concurrently with another one -- not even a callback on
a different object, as long as it is on the same client, which is why two
callbacks may share state without a lock between them. Do the slow part
elsewhere if that matters -- push the work onto a queue of your own and let
the callback return.

**An event's arguments are valid for the duration of the callback and no
longer.** They are COM objects the bridge minted for that one event, and it
releases them as soon as the callback returns -- whether it returned normally
or raised. Keeping one and using it later is a stale reference:

```ruby
kept = Queue.new
xl.ole_events.on('SheetChange') { |_sh, target| kept << target }
sheet.Range('C1').Value = 'hello'
target = kept.pop
sleep 1 # the callback has returned by now, so its arguments are released
begin
  target.Value
rescue => e
  puts "#{e.class}: #{e.message}"
end
puts "re-fetched instead: #{sheet.Range('C1').Value.inspect}"
```

```
WineOLE::RemoteError: WineOLE::StaleReferenceError: unknown handle 12884901895
re-fetched instead: "hello"
```

(The number is the released handle's id, so it differs from run to run.)

Take what you need inside the callback -- a name, an address, a value -- or
re-fetch the object afterwards by sheet name or address, as above.

**One event means the whole interface.** COM's `Advise` is per source
interface, not per event: subscribing to anything on `Application` puts the
sink on Excel's whole `AppEvents` interface, so every `Application` event
Excel raises reaches the bridge. The bridge forwards only the ones something
is subscribed to and drops the rest before minting any handles, so the cost
of the events you did not ask for is paid on the bridge and never on the
wire -- but it is paid, and a very chatty interface is worth knowing about.

**A callback that writes a cell fires `SheetChange` again.** Nothing here
guards against that, so guard it yourself:

```ruby
xl.ole_events.on('SheetChange') do |_sh, target|
  row = target.Row
  puts "callback saw row #{row}"
  sheet.Range("D#{row + 1}").Value = row + 1 if row < 3
end
sheet.Range('D1').Value = 1
```

```
callback saw row 1
callback saw row 2
callback saw row 3
```

This library deliberately never touches `Application.EnableEvents`. That
flag is global to the Excel instance -- shared with every macro in every open
workbook and with every other connection to the same instance -- so a library
that turned it off around a callback would be silently switching off someone
else's events. Set it yourself if you need it, and put it back.

**`args: false`** tells the bridge not to mint handles for an event's object
arguments. The callback is then called with no arguments at all (a block
written `|sheet, range|` gets `nil` for both), which is what you want for a
high-frequency event you only mean to count:

```ruby
xl.ole_events.on('SheetChange', args: false) do |*args|
  puts "changed; the callback was handed #{args.inspect}"
end
sheet.Range('E1').Value = 'x'
```

```
changed; the callback was handed []
```

```python
def counter(*args):
    print(f"changed; the callback was handed {list(args)!r}")

xl.ole_events.on('SheetChange', counter, args=False)
sheet.Range('E1').Value = 'x'
```

The bridge holds one such flag per event, so what goes on the wire is the
union of what the live callbacks asked for: register an ordinary callback
alongside an `args: false` one and the arguments come back for both, and
removing it drops them again. A callback that asked for `args: false` and
receives arguments anyway can ignore them; one that asked for arguments
always gets them.

While the flag is off nothing is minted, so nothing has to be given back:
those events cost no round trip to release them either, which is what keeps
the rate up on the high-frequency events this exists for.

**`on_error` receives what a callback raised**, and delivery continues --
both to the callbacks after it and to every later event:

```ruby
xl.ole_events.on_error { |err, frame| puts "#{frame['event']} raised #{err.class}: #{err.message}" }
xl.ole_events.on('SheetChange') { raise 'this callback is broken' }
xl.ole_events.on('SheetChange') { |_sh, t| puts "the next callback still ran, for #{t.Address}" }
sheet.Range('F1').Value = 'y'
sheet.Range('F2').Value = 'z'
```

```
SheetChange raised RuntimeError: this callback is broken
the next callback still ran, for $F$1
SheetChange raised RuntimeError: this callback is broken
the next callback still ran, for $F$2
```

```python
xl.ole_events.on_error(
    lambda err, frame: print(f"{frame['event']} raised {type(err).__name__}: {err}"))
```

Without an `on_error`, the same thing is reported on `$stderr` instead. The
raising callback is not unregistered: it will be called again on the next
event.

**Nothing a callback does can stop the events.** `on_error` sees whatever it
raised -- not only the `StandardError`s a bare `rescue` would catch, but a
`SystemStackError` out of a runaway recursion, or an `Exception` raised by
hand -- and the dispatcher goes on to the next
callback and the next event either way. A frame the bridge could not have
sent is reported the same way, rather than ending delivery inside the
machinery that parses it. This is deliberate: that thread *is* the delivery
mechanism, so losing it would be silent -- the callbacks would stay
registered, the bridge would stay advised, and every later event would be
minted, sent, queued and never delivered, with its COM objects never given
back. An `on_error` that itself raises is reported on `$stderr` and delivery
continues.

There is **one `on_error` per object**, and calling it again replaces the
handler rather than adding a second one -- unlike `on`, which appends. It
returns the `Events`, not a subscription, for that reason; `on_error { }`
is how you take one off. In Python the same handler is a callable of two
arguments, and `on_error(None)` is how you take one off -- Python has no
empty block to pass.

**`off` takes a callback back off**, by the subscription `on` returned or by
event name (which removes every callback for it):

```ruby
sub = xl.ole_events.on('SheetChange') { |_sh, _t| puts 'fired' }
sheet.Range('G1').Value = 1
sleep 1                         # let the callback run before taking it off
xl.ole_events.off(sub)          # or: xl.ole_events.off('SheetChange')
sheet.Range('G2').Value = 2
```

```
fired
```

```python
sub = xl.ole_events.on('SheetChange', lambda *args: print('fired'))
sheet.Range('G1').Value = 1
time.sleep(1)                    # let the callback run before taking it off
xl.ole_events.off(sub)           # or: xl.ole_events.off('SheetChange')
sheet.Range('G2').Value = 2
```

The second write raises `SheetChange` in Excel exactly as the first did;
nothing reaches the client because the last callback for it is gone, so the
subscription and the `Advise` under it are gone too.

`off(sub)` removes the registration `on` handed back and no other -- register
the same block twice and you have two of them, and taking one off leaves the
other firing. `off('SheetChange')` removes every callback for that event, and
`off` for an event nothing was ever registered for does nothing at all.

With the last callback on the object gone, the object stops taking part in
the connection -- and with the last callback on the connection, the
dispatcher thread stops too: an `ole_events` you have finished with costs
nothing, and one you never registered on costs nothing from the start.
`xl.ole_events.close` (Ruby) / `xl.ole_events.close()` (Python) is the bulk form -- every callback forgotten and every
subscription released, without having to remember the names -- and `on`
afterwards works exactly as the first one did.

**A dropped connection ends every subscription.** Closing the client -- or
losing the bridge -- unadvises everything that connection had advised, stops
the dispatcher and delivers nothing further. Nothing is queued, re-sent or
re-subscribed on reconnect: a new connection starts with no subscriptions,
and registering the callbacks again is what puts them back.

**Which object to subscribe on, for an MSForms control.** Excel hands out
three objects for one worksheet ActiveX control, and only one of them raises
events: the `OLEObject` is Excel's host (position, `LinkedCell`), its
`.Object` is the MSForms control -- that is the event source -- and *that*
object's own `.Object` is the raw control underneath, which has a `Caption`
but no `Name` and no events. Register on `sheet.OLEObjects('Go').Object.ole_events`;
the host answers `CONNECT_E_NOCONNECTION` and the raw control has no
connection points at all. A UserForm's controls are the same shape one layer
down: `ok.runtime.ole_events`, on the `Control` that `form.controls.add` or
`form.controls['OK']` hands back (`form.instance.Controls.Item('OK')` is the
same COM object reached the long way, but only `runtime`'s subscriptions are
closed by `form.unload`). Not `form.Controls`: that passes through to the
design-time control under `VBComponent.Designer`, which advises happily and
never fires. The UserForm itself is out of reach -- it exposes neither the
type information nor the connection points a sink needs under Wine -- so
form-level events such as `Initialize` and `QueryClose` cannot be received
by the client; handle those in VBA. The bridge finds the source interface
of these controls from their coclass type information (`IProvideClassInfo`),
since Wine has no proxy for the `IProvideClassInfo2` route it tries first.

## Remote bridges

`wineole-bridge` binds to `127.0.0.1` only by default, so it's unreachable
over the network unless you opt in. On the machine that will run the bridge:

```
WINEOLE_BIND=0.0.0.0 WINEOLE_TOKEN=<a secret> wine wineole-bridge.exe 47800
```

`WINEOLE_TOKEN` is optional but strongly recommended: without it, anything
that can reach that port can drive COM automation on that machine. The token
is only required from non-loopback connections, so this doesn't change
anything about local (`127.0.0.1`) use.

Then, from the client, pass `host:`/`host=` and `token:`/`token=` to
`Client.open`:

Python:

```python
client = wineole.Client.open(host='192.168.1.50', port=47800, token='<a secret>')
```

Ruby:

```ruby
client = WineOLE::Client.open(host: '192.168.1.50', port: 47800, token: '<a secret>')
```

`Client.open`'s auto-spawn fallback always launches the bridge on the
*local* machine -- it has no way to start one remotely. That's harmless when
the remote bridge is already running (the fast path just connects to it and
the spawn logic is never reached), but if it isn't running yet, the client
will spend the full `timeout` retrying a connection to a host nothing is
listening on yet, then raise, having pointlessly spawned a useless local
bridge in the meantime. Start the remote bridge yourself first, or pass a
`spawner:`/`spawner=` that raises instead of spawning locally, e.g.
`spawner: ->(port) { raise 'start the remote bridge manually' }`.

The wire protocol itself is plain, unencrypted TCP -- the token authenticates
but does not encrypt. Tunnel through SSH or a VPN if the network between the
client and the bridge isn't trusted.

## Prebuilt binaries

Prebuilt `wineole-bridge` binaries are cross-compiled for three Windows
architectures:

- `x86_64-pc-windows-gnu` and `i686-pc-windows-gnu`, built with the widely
  available `mingw-w64` toolchain (`apt install mingw-w64` on Debian/Ubuntu).
  These are exercised by this project's real-Excel integration tests on every
  change.
- `aarch64-pc-windows-gnullvm`, built with the
  [`llvm-mingw`](https://github.com/mstorsjo/llvm-mingw) toolchain instead
  (not packaged for apt — download a release tarball and put its `bin/` on
  `PATH`; on this project's development host it lives at
  `~/.local/share/llvm-mingw/`, so builds need
  `PATH="$HOME/.local/share/llvm-mingw/bin:$PATH"`.
  `wineole-bridge/.cargo/config.toml` points the target at
  `aarch64-w64-mingw32-clang` from it). **This binary has never been run.**
  This host (and every host this project has been developed on) has no way
  to execute an ARM64 Windows binary at all -- Wine here only supports
  x86/x86_64 execution, and there is no ARM64 CPU emulation layer (no
  `qemu-user`, no ARM64 Wine build) to bridge that gap. The binary is a
  structurally valid `PE32+ Aarch64` executable (verified with `file`) built
  from the same source as the other two targets, but nothing about its
  runtime correctness -- COM marshaling, struct layouts, calling conventions
  under the ARM64 ABI -- has been verified. Treat it as unverified until
  someone actually runs it on real ARM64 Wine.

On an unsupported host architecture both clients raise an explicit "no
prebuilt wineole-bridge binary for host architecture" error listing what is
available.

## Repository layout

`wineole-bridge/` is the shared Rust core: it builds the `wineole-bridge`
binaries and holds the prebuilt output in `wineole-bridge/dist/`. Each
language client lives in its own self-contained package directory under
`bindings/<lang>/` (`bindings/python/`, `bindings/ruby/`), so it can be built
and published independently. Because both RubyGems and setuptools reject
package file references that escape the package directory with `../`, each
binding instead contains a relative symlink back to `wineole-bridge/dist/`
(`bindings/python/wineole/dist` and `bindings/ruby/wineole-bridge-dist`) and
lists files through that symlink. These symlinks are load-bearing, not
decorative: if one is deleted or replaced with a real (empty) directory, the
gem or wheel will still build and install successfully, then fail at runtime
with no bridge binary present.
