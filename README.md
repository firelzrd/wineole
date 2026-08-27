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
that has to actually be called. See `bindings/ruby/examples/` and
`bindings/python/examples/` for a larger, visual demo of both.

### Meta-methods and the `ole_` prefix

A `Proxy` forwards any name it doesn't recognize straight to the remote COM
object, so its own bookkeeping methods are named with an `ole_` prefix --
`ole_handle`, `ole_session_id`, `ole_release`, `ole_const_load`, and
`ole_created?`/`ole_created` (below) -- to keep them out of the way of a
real COM member that happens to share the plain name. `invoke` is the one
deliberate exception: it's kept bare and public as an explicit escape hatch
for the rare case where a COM object really does define a member called
e.g. `ole_handle` -- `proxy.invoke('ole_handle', [], {})` reaches it
directly, bypassing the local method entirely. This mirrors real Ruby
`WIN32OLE`'s own `ole_*`-prefixed introspection methods and its own bare,
public `invoke`.

`create`/`connect`/`connect_or_create` set `ole_created?`/`ole_created` on
the `Proxy` they return: `true`/`True` if a new instance was created,
`false`/`False` if an existing one was attached to, and `nil`/`None` for
any object derived from another (`xl.Worksheets`, `xl.Workbooks.Add`, ...)
-- attach-vs-create isn't a meaningful question for those. Handy for
deciding whether to `Quit` on the way out:

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

`Client.open`/`wineole`'s equivalents reuse an already-running bridge on the
given port, or spawn one and wait for it to come up -- either language's
client can talk to a bridge the other one started. `client.close()`/
`client.close` (or the implicit default's own eventual cleanup) releases the
session; the bridge automatically frees everything the connection owned (COM
objects included) once it sees the connection drop, and shuts down after 30
minutes of no active connections.

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
  `PATH`; `wineole-bridge/.cargo/config.toml` points the target at
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
  It is also currently behind the other two: it predates the
  `connect_or_create` RPC and will reject that call, so
  `connect_or_create`/`ole_created?`/`ole_created` do not work on ARM64
  until someone rebuilds it with `llvm-mingw` on `PATH`.

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
