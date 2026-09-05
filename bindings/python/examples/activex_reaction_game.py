#!/usr/bin/env python3
"""A reaction-time game played entirely inside Excel: nine ActiveX
CommandButtons in a 3x3 grid. One lights up red at a time; click it before
the next one lights, and Python scores how fast you were. The Python twin
of bindings/ruby/examples/activex_reaction_game.rb.

Every state transition here runs INSIDE a Click callback -- there is no
polling loop and no timer thread. Lighting the next target, scoring a hit,
flashing a miss, ending the round: all of it happens on the connection's
one dispatcher thread, in the order Excel delivered the clicks. That
thread is shared with every other callback on this connection, so the
brief sleep on a miss (a visible flash, not just a colour swap) delays
whatever click follows it by that much -- acceptable for a demo, worth
knowing for a real one.

The next target is never the square just hit: an unchanged red square
reads as "nothing happened yet" for a beat, costing the player real
reaction time on exactly the round that would otherwise be their fastest.

Run: python3 bindings/python/examples/activex_reaction_game.py
"""
import os
import queue
import random
import sys
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.msoffice import Excel, Color
import wineole

ROUNDS = 10
GRID = 3
IDLE = '#D9D9D9'
TARGET = '#FF3333'
MISS = '#FFA500'
MISS_FLASH_SECONDS = 0.25

with Excel.run('create') as xl:
    xl.show()
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
    moles = []
    for i in range(GRID * GRID):
        row, col = divmod(i, GRID)
        moles.append(sheet.activex.add(
            'command_button', name=f'Mole{i}', caption='',
            left=base_left + col * (size + gap), top=base_top + row * (size + gap),
            width=size, height=size))

    start_button = sheet.activex.add(
        'command_button', name='StartButton', caption='Start',
        left=base_left, top=base_top + GRID * (size + gap) + 10, width=size, height=30)
    quit_button = sheet.activex.add(
        'command_button', name='QuitButton', caption='Quit',
        left=base_left + size + gap, top=base_top + GRID * (size + gap) + 10,
        width=size, height=30)

    def reset_grid():
        for m in moles:
            m.BackColor = Color.parse(IDLE)

    reset_grid()

    state = {'playing': False, 'round': 0, 'score': 0, 'target': None,
             'lit_at': None, 'best': None}
    done = queue.Queue()

    def next_round():
        state['round'] += 1
        if state['round'] > ROUNDS:
            state['playing'] = False
            reset_grid()
            sheet['A8'] = f"Done -- score {state['score']}/{ROUNDS}. " \
                          'Click Start to play again, or Quit.'
            return
        round_cell.write(state['round'])
        reset_grid()
        previous_target = state['target']
        target = random.randint(0, len(moles) - 2)
        if previous_target is not None and target >= previous_target:
            target += 1
        state['target'] = target
        state['lit_at'] = time.monotonic()
        moles[target].BackColor = Color.parse(TARGET)

    def on_start(*_args):
        state.update(playing=True, round=0, score=0)
        sheet['A8'] = ''
        score_cell.write(0)
        last_cell.write('')
        next_round()
        print('[Python] Game started.')

    def make_mole_handler(i):
        def on_click(*_args):
            if not state['playing']:
                return

            if i == state['target']:
                elapsed = time.monotonic() - state['lit_at']
                state['score'] += 1
                if state['best'] is None or elapsed < state['best']:
                    state['best'] = elapsed
                score_cell.write(state['score'])
                last_cell.write(round(elapsed, 3))
                best_cell.write(round(state['best'], 3))
                print(f"[Python] Hit #{i} in {elapsed:.3f}s (score {state['score']})")
                next_round()
            else:
                print(f"[Python] Miss -- clicked #{i}, target was #{state['target']}")
                moles[i].BackColor = Color.parse(MISS)
                time.sleep(MISS_FLASH_SECONDS)
                if i != state['target']:
                    moles[i].BackColor = Color.parse(IDLE)

        return on_click

    def on_quit(*_args):
        done.put(True)

    start_button.on('Click', on_start)
    mole_handlers = [make_mole_handler(i) for i in range(len(moles))]
    for mole, handler in zip(moles, mole_handlers):
        mole.on('Click', handler)
    quit_button.on('Click', on_quit)

    print('Excel is up. Click Start, then hit the red square each round. Quit when done.')
    done.get()
    print('[Python] Quit clicked -- shutting down.')

    for control in [*moles, start_button, quit_button]:
        control.off('Click')

wineole.close()
