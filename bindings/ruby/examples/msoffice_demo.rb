#!/usr/bin/env ruby
# frozen_string_literal: true

# Visual demo, not a test: a wineole port of lib/ruby/msoffice.rb's own
# self-test block ($0 == __FILE__) -- caption several worksheets, fill a
# new sheet solid with a block character, cycle through all sheets with
# pauses, then hammer the last sheet with 10 threads writing random
# colored cells simultaneously for a few seconds.
#
# Unlike the original, worker threads stop via a cooperative flag rather
# than Thread#kill. msoffice.rb's raw WIN32OLE calls are independent Win32
# API calls per OS thread, so killing one mid-call is harmless there. Here
# all threads share one TCP connection with a line-framed JSON protocol
# (WineOLE::Client#call), serialized by one Mutex -- killing a thread while
# it holds that mutex mid-write can leave a half-written line on the wire,
# corrupting the stream for every other thread and the main thread's own
# later calls. Confirmed empirically: an early run using Thread#kill left
# Excel actually quitting correctly, but the *response* to that Quit call
# came back truncated (JSON::ParserError), because a killed thread had
# clipped a previous write mid-flight.
#
# Run: ruby bindings/ruby/examples/msoffice_demo.rb

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'wineole'

xl = WineOLE.create('Excel.Application')

begin
  xl.Visible = true
  xl.DisplayAlerts = false
  xl.Workbooks.Add

  (3 - xl.Worksheets.Count).times { xl.Worksheets.Add }

  xl.Worksheets[1].Range('A1').Value = 'This is the first worksheet'
  xl.Worksheets[2].Range('A1').Value = 'This is the 2nd worksheet'
  xl.Worksheets[xl.Worksheets.Count].Range('A1').Value = 'This is the last worksheet'

  block_sheet = xl.Worksheets.Add
  block_sheet.Range('A1:AN20').Value = '□'

  xl.Worksheets.Count.times { |i| xl.Worksheets[i + 1].Cells.ColumnWidth = 2 }

  xl.Worksheets[1].Select
  sleep 3
  xl.Worksheets[2].Select
  sleep 3
  xl.Worksheets[3].Select
  sleep 3
  block_sheet.Select

  stop = false
  threads = Array.new(10) do
    Thread.new do
      until stop
        100.times do
          begin
            cell = block_sheet.Cells(rand(20) + 1, rand(40) + 1)
            cell.Interior.ColorIndex = rand(48)
            cell.Value = rand(100)
          rescue WineOLE::RemoteError
            # msoffice.rb's own demo swallows the same class of transient
            # COM error the same way -- ColorIndex 0 is out of Excel's
            # valid 1..56 palette range and occasionally gets rolled here.
          end
        end
      end
    end
  end
  sleep 10
  stop = true
  threads.each(&:join)

  puts 'msoffice.rb-style demo finished.'
ensure
  begin
    xl.Quit
  ensure
    WineOLE.close
  end
end
