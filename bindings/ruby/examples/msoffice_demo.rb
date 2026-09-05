#!/usr/bin/env ruby
# frozen_string_literal: true

# Visual demo, not a test: a wineole port of lib/ruby/msoffice.rb's own
# self-test block -- caption several worksheets, fill a new sheet solid with
# a block character, cycle through them with pauses, then hammer the last
# one with 10 threads writing random coloured cells for a few seconds.
#
# It used to be written against raw COM, and what went away in the port was
# all plumbing: the nested begin/ensure around Quit, the
# xl.Worksheets[n].Range('A1').Value chains, and the Interior.ColorIndex
# arithmetic (a bare 1..56 integer, so rand(48) sometimes rolled a 0 and had
# to be rescued). Excel 2003 still snaps '#RRGGBB' down to its own 56-colour
# palette, so the screen looks the same -- the wrapper removed the
# arithmetic, not the palette.
#
# Checked rather than assumed: Excel.run quits the COM Application but never
# touches the Client, so WineOLE.close at the end is still this script's job.
#
# The worker threads stop via a cooperative flag rather than Thread#kill,
# and wrapping the COM calls does not change that: all ten share one TCP
# connection behind one Mutex, so killing a thread mid-write leaves half a
# line on the wire and corrupts the stream for every other thread and for
# the main thread's later calls. Observed on an earlier version of this
# demo -- Excel quit correctly, but the *response* to that Quit came back
# truncated (JSON::ParserError), clipped by a thread killed mid-flight.
#
# Run: ruby bindings/ruby/examples/msoffice_demo.rb

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'wineole/msoffice'

# The alias the README shows and tells readers to write themselves --
# WineOLE::MSOffice claims nothing in the root namespace on its own.
MSOffice = WineOLE::MSOffice

MSOffice::Excel.run(:create) do |xl|
  xl.show

  xl.no_alert do
    xl['[:new]']
    (3 - xl.Worksheets.Count).times { xl[':new!'] }

    xl['1!A1'] = 'This is the first worksheet'
    xl['2!A1'] = 'This is the 2nd worksheet'
    xl[':last!A1'] = 'This is the last worksheet'

    block_sheet = xl[':new!']
    block_sheet['A1:AN20'].fill('□')

    xl['[]'].each_sheet { |sheet| sheet.Cells.ColumnWidth = 2 }

    [xl['1!'], xl['2!'], xl['3!']].each do |sheet|
      sheet.Select
      sleep 3
    end
    block_sheet.Select

    stop = false
    threads = Array.new(10) do
      Thread.new do
        until stop
          100.times do
            # format returns the Range, so the colour and the value are
            # one expression.
            block_sheet[rand(20) + 1, rand(40) + 1]
              .format(background: Kernel.format('#%06X', rand(0x1000000)))
              .write(rand(100))
          end
        end
      end
    end
    sleep 10
    stop = true
    threads.each(&:join)
  end

  puts 'msoffice_demo.rb finished.'
end

WineOLE.close
