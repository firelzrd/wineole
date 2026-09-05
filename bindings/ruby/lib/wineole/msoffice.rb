# The bundled Microsoft Office wrapper: Address, Paths, Range, Sheet, Book,
# Excel, and the Controls and Forms wrappers, all under WineOLE::MSOffice.
#
# Deliberately NOT required from lib/wineole.rb. The core layer is a
# general-purpose COM bridge; this is Office-specific, and it is where a
# Word and a PowerPoint wrapper will land too, so it will only grow.
# Someone who wants the bridge should not have to carry an Excel wrapper to
# get it.
#
# Defines no top-level constant. An earlier draft aliased WineOLE::MSOffice
# to a root-level MSOffice for code written against the older msoffice.rb;
# reaching into the root namespace from a library is intrusive enough that
# the alias was not worth the two failure modes it brought with it (a
# silent no-op when something else already defined MSOffice, and a test
# whose result depended on process-wide load order).
require_relative 'msoffice/address'
require_relative 'msoffice/vba'
require_relative 'msoffice/vba_block'
require_relative 'msoffice/vba_api'
require_relative 'msoffice/color'
require_relative 'msoffice/format'
require_relative 'msoffice/paths'
require_relative 'msoffice/range'
require_relative 'msoffice/sheet'
require_relative 'msoffice/controls'
require_relative 'msoffice/forms'
require_relative 'msoffice/book'
require_relative 'msoffice/excel'
