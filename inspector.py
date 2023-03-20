import servant

import openpyxl
from openpyxl.utils import column_index_from_string

from pathlib import Path
import subprocess
import pyautogui
import time


class Inspector:
    def __init__(self, up_spreadsheet, lo_spreadsheet):
        self.file_up = up_spreadsheet
        self.file_lo = lo_spreadsheet

    @property
    def refund_up(self):
        wb = openpyxl.load_workbook(self.file_up, data_only=True)
        ws = wb.worksheets[0]
        ref_val = ws.cell(21, 7).value

        if ref_val is None:
            print(f'Formulas in file {self.file_up} was not calculated yet. Opening installed spreadsheet program')
            resave_xlsx(self.file_up)
            wb = openpyxl.load_workbook(self.file_up, data_only=True)
            ws = wb.worksheets[0]
            return ws.cell(21, 7).value
        else:
            return ref_val

    @property
    def refund_lo(self):
        wb = servant.unlock(self.file_lo, 13881744, data_only=True)
        ws = wb.worksheets[-2]
        q = self.quarter
        if q:
            ref_val = ws.cell(q + 2, column_index_from_string('R')).value
            if ref_val is None:
                print(f'Formulas in file {self.file_lo} was not calculated yet. Opening installed spreadsheet program')
                resave_xlsx(self.file_lo)
                wb = servant.unlock(self.file_lo, 13881744, data_only=True)
                ws = wb.worksheets[-2]
                return ws.cell(q + 2, column_index_from_string('R')).value
            else:
                return ref_val
        else:
            print(f'Quarter is not defined. Check file {self.file_up}, cell D4')

    @property
    def quarter(self):
        wb = openpyxl.load_workbook(self.file_up)
        ws = wb.worksheets[0]
        quarter = ws.cell(6, 4).value

        try:
            return int(quarter)
        except ValueError:
            print(f'Quarter is not integer type. Fix the value in {self.file_up} in cell D6')
            return False

    def refund_are_equal(self):
        rup = self.refund_up
        rlo = self.refund_lo
        if rup == rlo:
            return True
        else:
            print(f'RUP: {rup}, RLO: {rlo}, difference: {rup - rlo}')
            return False


file_up = 'temp-up.xlsx'
file_lo = 'temp.xlsx'


def resave_xlsx(file):

    if servant.is_tool('libreoffice'):
        xlsx_tool = 'libreoffice'
    elif servant.is_tool('ms excel'):
        xlsx_tool = 'ms excel'
    elif servant.is_tool('openoffice'):
        xlsx_tool = 'openoffice'
    else:
        print('No xlsx processor installed. Please install MS Excel, LibreOffice or OpenOffice')
        return 0

    # Open xlsx editing software
    cmd = [xlsx_tool, file]
    subprocess.Popen(cmd)

    # Wait until spreadsheet is loaded
    time.sleep(3)

    # Switch focus to edit software
    cmd = ['wmctrl', '-a', xlsx_tool]
    subprocess.Popen(cmd)

    # Save as
    pyautogui.hotkey('ctrl', 'shift', 's')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('right')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

    # Close the program
    pyautogui.hotkey('alt', 'f4')
    time.sleep(0.5)


if all((Path(file_up).exists(), Path(file_lo).exists())):
    inspector = Inspector(file_up, file_lo)
    print(inspector.refund_are_equal())
else:
    print('File does not exist')
