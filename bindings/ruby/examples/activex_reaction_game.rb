#!/usr/bin/env ruby
# frozen_string_literal: true

# A reaction-time game played entirely inside Excel: nine ActiveX
# CommandButtons in a 3x3 grid. One lights up red at a time; click it
# before the next one lights, and Ruby scores how fast you were.
#
# Every state transition here runs INSIDE a Click callback -- there is no
# polling loop and no timer thread. Lighting the next target, scoring a
# hit, flashing a miss, ending the round: all of it happens on the
# connection's one dispatcher thread, in the order Excel delivered the
# clicks (see wineole/events.rb). That thread is shared with every other
# callback on this connection, so the brief sleep on a miss (a visible
# flash, not just a colour swap) delays whatever click follows it by that
# much -- acceptable for a demo, worth knowing for a real one.
#
# Run: ruby bindings/ruby/examples/activex_reaction_game.rb

$LOAD_PATH.unshift(File.expand_path('../lib', __dir__))
require 'wineole/msoffice'

MSOffice = WineOLE::MSOffice
Color = MSOffice::Color

ROUNDS = 10
GRID = 3
IDLE = '#D9D9D9'
TARGET = '#FF3333'
MISS = '#FFA500'
MISS_FLASH_SECONDS = 0.25

MSOffice::Excel.run(:create) do |xl|
  xl.show
  xl['[:new]']
  sheet = xl[':first!']
  sheet.ole.Columns('A').ColumnWidth = 7
  sheet.ole.Columns('B:H').ColumnWidth = 3

  sheet['A1'] = 'Reaction game -- Start, then click the RED square as fast as you can'
  sheet['A3'] = 'Round'
  sheet['A4'] = 'Score'
  sheet['A5'] = 'Last (s)'
  sheet['A6'] = 'Best (s)'
  round_cell = sheet['B3']
  score_cell = sheet['B4']
  last_cell = sheet['B5']
  best_cell = sheet['B6']

  # Grid buttons live at absolute points, independent of the narrow A:H
  # columns above, so the two layouts never fight over the same space.
  base_left, base_top, size, gap = 260, 40, 60, 12
  moles = Array.new(GRID * GRID) do |i|
    row, col = i.divmod(GRID)
    sheet.activex.add(
      :command_button, name: "Mole#{i}", caption: '',
      left: base_left + col * (size + gap), top: base_top + row * (size + gap),
      width: size, height: size
    )
  end

  start_button = sheet.activex.add(
    :command_button, name: 'StartButton', caption: 'Start',
    left: base_left, top: base_top + GRID * (size + gap) + 10, width: size, height: 30
  )
  quit_button = sheet.activex.add(
    :command_button, name: 'QuitButton', caption: 'Quit',
    left: base_left + size + gap, top: base_top + GRID * (size + gap) + 10, width: size, height: 30
  )

  moles.each { |m| m.BackColor = Color[IDLE] }

  playing = false
  round = 0
  score = 0
  target = nil
  lit_at = nil
  best = nil
  done = Queue.new

  reset_grid = -> { moles.each { |m| m.BackColor = Color[IDLE] } }

  next_round = lambda do
    round += 1
    if round > ROUNDS
      playing = false
      reset_grid.call
      sheet['A8'] = "Done -- score #{score}/#{ROUNDS}. Click Start to play again, or Quit."
      next
    end
    round_cell.write(round)
    reset_grid.call
    # Never repeats the square just hit -- an unchanged red square reads as
    # "nothing happened yet" for a beat, costing the player real reaction
    # time on exactly the round that would otherwise be their fastest.
    previous_target = target
    target = rand(moles.length - 1)
    target += 1 if !previous_target.nil? && target >= previous_target
    lit_at = Time.now
    moles[target].BackColor = Color[TARGET]
  end

  start_button.on('Click') do
    playing = true
    round = 0
    score = 0
    sheet['A8'] = ''
    score_cell.write(0)
    last_cell.write('')
    next_round.call
    puts '[Ruby] Game started.'
  end

  moles.each_with_index do |mole, i|
    mole.on('Click') do
      next unless playing

      if i == target
        elapsed = Time.now - lit_at
        score += 1
        best = elapsed if best.nil? || elapsed < best
        score_cell.write(score)
        last_cell.write(elapsed.round(3))
        best_cell.write(best.round(3))
        puts "[Ruby] Hit ##{i} in #{elapsed.round(3)}s (score #{score})"
        next_round.call
      else
        puts "[Ruby] Miss -- clicked ##{i}, target was ##{target}"
        mole.BackColor = Color[MISS]
        sleep MISS_FLASH_SECONDS
        mole.BackColor = Color[IDLE] unless i == target
      end
    end
  end

  quit_button.on('Click') { done << true }

  puts 'Excel is up. Click Start, then hit the red square each round. Quit when done.'
  done.pop
  puts '[Ruby] Quit clicked -- shutting down.'

  (moles + [start_button, quit_button]).each { |c| c.off('Click') }
end

WineOLE.close
