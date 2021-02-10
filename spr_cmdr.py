#!/usr/bin/env python3
'''
Spreadsheet Commander
---------------------
A basic spreadsheet transfer program

If the program acts weird while controlling the keyboard,
smash the mouse cursor into the corner of the screen to abort.

Usage:
======
python3 spr_cmdr.py [OPTIONS]

Options:
    --csv CSV_FILE                      Parse as a CSV (Comma Separated Values) file.
    -e, -x, --xl, --excel EXCEL_FILE    Parse as an Excel-compatible file (see --excel-sheet-*).
    --excel-sheet-num SHEET_NUM         Choose sheet number (zero-indexed) to parse in Excel-compatible files. Defaults to the first sheet.
    --excel-sheet-name SHEET_NAME       Choose sheet name to parse in Excel-compatible files. Has priority over --excel-sheet-num.
    -c, --col-head-idx INDEX            Choose which column is the header column (which will be ignored). Zero-indexed, default None.
    -r, --row-head-idx INDEX            Choose which row is the header row (which will be ignored). Zero-indexed, default None.
    -t, --timer SECONDS                 Determines how many seconds the program waits to start transfer. Default 3.
    -n, --nan                           Pass this option to enable NaN filtering.
'''

import time

import pandas as pd
import numpy as np
import pyautogui
import click

pyautogui.PAUSE = 0.2

@click.command()
@click.option('--csv', type=click.Path(exists=True, dir_okay=False), default=None, help='Parse as a CSV (Comma Separated Values) file')
@click.option('--excel', '--xl', '-e', '-x', type=click.Path(exists=True, dir_okay=False), default=None, help='Parse as an Excel-compatible file (see --excel-sheet-*)')
@click.option('--excel-sheet-num', default=0, help='Choose sheet number (zero-indexed) to parse in Excel-compatible files. Defaults to the first sheet.')
@click.option('--excel-sheet-name', type=str, help='Choose sheet name to parse in Excel-compatible files. Has priority over --excel-sheet-num.')
@click.option('--col-head-idx', '-c', type=int, help='Choose which column is the header column (which will be ignored). Zero-indexed, default None.')
@click.option('--row-head-idx', '-r', type=int, help='Choose which row is the header row (which will be ignored). Zero-indexed, default None.')
@click.option('--timer', '-t', default=3, help='Determines how many seconds the program waits to start transfer. Default 3.')
@click.option('--nan', '-n', is_flag=True, help='Pass this option to enable NaN filtering.')
def spr_cmdr(csv, excel, excel_sheet_num, excel_sheet_name, col_head_idx, row_head_idx, timer, nan):
    data: pd.DataFrame = None
    if csv is not None:
        data = pd.read_csv(csv, index_col=col_head_idx, header=row_head_idx)
    elif excel is not None:
        if excel_sheet_name is not None:
            data = pd.read_excel(excel, sheet_name=excel_sheet_name, index_col=col_head_idx, header=row_head_idx)
        else:
            data = pd.read_excel(excel, sheet_name=excel_sheet_num, index_col=col_head_idx, header=row_head_idx)
    else:
        click.echo('Error: No input file provided!\nSee --help for usage.', err=True)
        return
    click.echo('Spreadsheet Loaded!')
    click.pause('Press any key to start transfer timer.')
    time.sleep(timer)
    click.echo('Running!')
    for _, col in data.items():
        ln = 0
        for val in col:
            if pd.notna(val):
                pyautogui.write(str(val))
            ln += 1
            pyautogui.press('down')
        pyautogui.press('up', presses=ln, interval=0.2)
        pyautogui.press('right')
    click.echo('Done!')



if __name__ == '__main__':
    spr_cmdr()

