# Changelog

All notable changes to this project are documented here. Format loosely
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [0.2.0] - 2026-09-06

### Added

- **COM events**: the bridge advises COM connection points and streams
  callbacks back over the wire, so a client's `on(event) { ... }` /
  `on(event, callback)` runs on a per-connection dispatcher thread as
  events arrive. Available in both the Ruby and the Python client.
- **MSOffice wrapper**, shipped for both Ruby and Python:
  - `Excel` (`create`/`connect`/`connect_or_create`/`run`), `Book`,
    `Sheet`, `Range` and an addressing DSL (`xl['Sheet1!A1:B10']`) that
    resolves workbooks, worksheets and ranges from one string.
  - `Range#write`/`#fill`/`#format`, `Format` (fonts, borders, number
    formats) and `Color` (`'#RRGGBB'` &lt;-&gt; Excel's BGR integer).
  - `Paths`, converting between Wine-side and host-side file paths.
  - A `Passthrough` base that routes unknown calls straight to the
    underlying COM object.
- **VBA API and CLI**: `VBA` (registry-based macro security state,
  ANSI-codepage detection), `VBABlock` (marker-delimited managed code
  blocks), `BookVBA`/`SheetVBA` (`write`, `remove`, `add_component`,
  `import`/`export`), and a `wineole-vba` console script for both
  bindings.
- **Controls and Forms**: `Controls`/`Control` and the three placement
  families (`FormControls`, `ActiveXControls`, `UserFormControls`), plus
  `Forms`/`Form` for UserForms (`show`/`hide`/`unload`, VBA handler
  injection, and Click events reaching Ruby/Python directly for ActiveX
  and UserForm controls).
- Example programs: an ActiveX button demo and a reaction-time minigame
  (`bindings/*/examples/activex_reaction_game.*`) driving Excel entirely
  from COM event callbacks, alongside the existing worksheet demo.
- A `wineole-vba` executable entry for the Ruby gem and the Python
  package (`[project.scripts]` / `spec.executables`).

### Changed

- Expanded `wineole-bridge`'s session/sink/value layer substantially to
  support COM event advising and delivery.
- README rewritten to document the Office wrapper, the VBA layer,
  Controls/Forms and COM events for both bindings, side by side.

## [0.1.0] - 2026-08-28

### Added

- Initial `wineole-bridge` (Rust, cross-compiled for Wine): a JSON Lines
  TCP protocol for driving Win32OLE/COM automation of Windows
  applications running under Wine.
- Initial Ruby client (`wineole` gem): `Client`, `Proxy`, `Dispatcher`,
  instance ownership and cleanup, and the `msoffice_demo.rb` example.
- Initial Python client scaffolding (`wineole` package).
