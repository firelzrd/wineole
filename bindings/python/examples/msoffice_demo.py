#!/usr/bin/env python3
"""Visual demo, not a test: caption several worksheets, fill a new sheet
solid with '#' characters, cycle through them with pauses, then hammer
the last one with 10 threads writing random coloured cells for a few
seconds. The Python twin of bindings/ruby/examples/msoffice_demo.rb.

It used to be written against raw COM, and what went away in the port was
all plumbing: the try/finally around Quit, the
xl.Worksheets()[n].Range('A1').Value chains, and the Interior().ColorIndex
arithmetic (a bare 1..56 integer, so a random 0..47 sometimes rolled a 0 and
had to be caught). Excel 2003 still snaps '#RRGGBB' down to its own
56-colour palette, so the screen looks the same -- the wrapper removed the
arithmetic, not the palette.

Checked rather than assumed: Excel.run releases the COM Application but
never touches the Client, so wineole.close() at the end is still this
script's job.

The worker threads stop via a cooperative flag. Python has no safe way to
force-kill a running thread, and that is the right shape here anyway: all
ten share one TCP connection behind one lock, so ending a thread mid-write
would leave half a line on the wire and corrupt the stream for every other
thread and for the main thread's later calls.

Run: python3 bindings/python/examples/msoffice_demo.py
"""
import os
import random
import sys
import threading
import time

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from wineole.msoffice import Excel
import wineole

try:
    with Excel.run('create') as xl:
        xl.show()

        with xl.no_alert():
            xl['[:new]']
            for _ in range(3 - xl.Worksheets().Count()):
                xl[':new!']

            xl['1!A1'] = 'This is the first worksheet'
            xl['2!A1'] = 'This is the 2nd worksheet'
            xl[':last!A1'] = 'This is the last worksheet'

            block_sheet = xl[':new!']
            block_sheet['A1:AN20'].fill('#')

            for sheet in xl['[]'].sheets():
                sheet.Cells().ColumnWidth = 2

            for address in ('1!', '2!', '3!'):
                xl[address].Select()
                time.sleep(3)
            block_sheet.Select()

            stop = threading.Event()

            def hammer():
                while not stop.is_set():
                    for _ in range(100):
                        # format returns the Range, so the colour and the value
                        # are one expression.
                        block_sheet[random.randint(1, 20), random.randint(1, 40)] \
                            .format(background='#%06X' % random.randint(0, 0xFFFFFF)) \
                            .write(random.randint(0, 99))

            threads = [threading.Thread(target=hammer) for _ in range(10)]
            for t in threads:
                t.start()
            time.sleep(10)
            stop.set()
            for t in threads:
                t.join()

        print('msoffice_demo.py finished.')
finally:
    wineole.close()
