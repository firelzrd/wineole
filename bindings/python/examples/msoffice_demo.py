#!/usr/bin/env python3
"""Visual demo, not a test: a wineole port of lib/ruby/msoffice.rb's own
self-test block ($0 == __FILE__) -- caption several worksheets, fill a
new sheet solid with a block character, cycle through all sheets with
pauses, then hammer the last sheet with 10 threads writing random
colored cells simultaneously for a few seconds.

Unlike the Ruby port, this uses a cooperative threading.Event to stop the
worker threads instead of Ruby's Thread#kill -- Python has no safe way to
force-kill a running thread, so a stop flag the threads check between
batches is the correct (not just stylistic) way to end them.

Run: python3 bindings/python/examples/msoffice_demo.py
"""
import os
import random
import sys
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

import wineole

xl = wineole.create('Excel.Application')

try:
    xl.Visible = True
    xl.DisplayAlerts = False
    xl.Workbooks().Add()

    for _ in range(3 - xl.Worksheets().Count()):
        xl.Worksheets().Add()

    xl.Worksheets()[1].Range('A1').Value = 'This is the first worksheet'
    xl.Worksheets()[2].Range('A1').Value = 'This is the 2nd worksheet'
    xl.Worksheets()[xl.Worksheets().Count()].Range('A1').Value = 'This is the last worksheet'

    block_sheet = xl.Worksheets().Add()
    block_sheet.Range('A1:AN20').Value = '□'

    for i in range(xl.Worksheets().Count()):
        xl.Worksheets()[i + 1].Cells().ColumnWidth = 2

    xl.Worksheets()[1].Select()
    time.sleep(3)
    xl.Worksheets()[2].Select()
    time.sleep(3)
    xl.Worksheets()[3].Select()
    time.sleep(3)
    block_sheet.Select()

    stop = threading.Event()

    def hammer():
        while not stop.is_set():
            for _ in range(100):
                try:
                    cell = block_sheet.Cells(random.randint(1, 20), random.randint(1, 40))
                    cell.Interior().ColorIndex = random.randint(0, 47)
                    cell.Value = random.randint(0, 99)
                except wineole.RemoteError:
                    # msoffice.rb's own demo swallows the same class of
                    # transient COM error the same way -- ColorIndex 0 is
                    # out of Excel's valid 1..56 palette range and
                    # occasionally gets rolled here.
                    pass

    threads = [threading.Thread(target=hammer) for _ in range(10)]
    for t in threads:
        t.start()
    time.sleep(10)
    stop.set()
    for t in threads:
        t.join()

    print('msoffice.rb-style demo finished.')
finally:
    try:
        xl.Quit()
    finally:
        wineole.close()
