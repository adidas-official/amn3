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
        self.wb_up, self.wb_lo = self.is_readable()
        self.q = self.quarter

    def is_readable(self):
        wb_up = openpyxl.load_workbook(self.file_up, data_only=True)
        ws = wb_up.worksheets[0]

        if ws.cell(21, 7).value is None:
            resave_xlsx(self.file_up)
            wb_up = openpyxl.load_workbook(self.file_up, data_only=True)

        wb_lo = servant.unlock(self.file_lo, '13881744', data_only=True)
        ws = wb_lo.worksheets[-2]

        if ws.cell(3, 2).value is None:
            resave_xlsx(self.file_lo)
            wb_lo = servant.unlock(self.file_lo, '13881744', data_only=True)

        return wb_up, wb_lo

    @property
    def refund_up(self):

        ws = self.wb_up.worksheets[0]
        ref_val = ws.cell(21, 7).value

        return ref_val

    @property
    def refund_lo(self):
        ws = self.wb_lo.worksheets[-2]

        if self.q:
            ref_val = ws.cell(self.q + 2, column_index_from_string('R')).value
            return ref_val
        else:
            print(f'Quarter is not defined. Check file {self.file_up}, cell D4')

    @property
    def quarter(self):
        ws = self.wb_up.worksheets[0]
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
            print('Refunds are not equal')
            print(f'RUP: {rup}, RLO: {rlo}, difference: {rup - rlo}')
            return False

    def sum_month_refund(self, month):
        ws = self.wb_up.worksheets[1]

        row = 13
        total = 0

        while True:
            if not ws.cell(row, 2).value:
                break

            for i in range(3):
                val = ws.cell(row, 12 + (month * 5) + i).value
                if val:
                    total += val
            row += 1
        return total

    @property
    def sums_of_refunds(self):
        months = [self.sum_month_refund(i) for i in range(3)]
        return months

    def refund_for_month_lo(self, month):
        month_index = self.q * 3 - 2 + month
        ws = self.wb_lo.worksheets[-2]
        return ws.cell(13, month_index + 1).value

    @property
    def sums_of_refunds_lo(self):
        return [self.refund_for_month_lo(month) for month in range(3)]

    @property
    def faulty_months(self):
        if not self.refund_are_equal():
            faulty_months = []
            refunds_up = self.sums_of_refunds
            refunds_lo = self.sums_of_refunds_lo

            for month_index, month in enumerate(zip(refunds_up, refunds_lo)):
                if month[0] != month[1]:
                    faulty_months.append(month_index)

            return faulty_months
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
    print(inspector.faulty_months)
    # print(inspector.refund_are_equal())
else:
    print('File does not exist')
