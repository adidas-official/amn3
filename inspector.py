import servant
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path


class Inspector:
    def __init__(self, up_spreadsheet, lo_spreadsheet):
        self.wb_up = openpyxl.load_workbook(up_spreadsheet, data_only=True)
        self.wb_lo = servant.unlock(lo_spreadsheet, '13881744', data_only=True)

    @property
    def refund_up(self):
        ws = self.wb_up.worksheets[0]
        return ws.cell(21, 7).value

    @property
    def refund_lo(self):
        ws = self.wb_lo.worksheets[-2]
        q = self.quarter

        if type(q) != int:
            if q.isnumeric():
                q = int(q)
            else:
                return 'Quarter is not integer'
        else:
            return ws.cell(q, column_index_from_string('R')).value

    @property
    def quarter(self):
        ws = self.wb_up.worksheets[0]
        return ws.cell(6, 4).value

    def refund_are_equal(self):
        if self.refund_up == self.refund_lo:
            return True
        return False


file_up = 'tables/jmenny_seznam_2022_09_27 Bereko.xlsx'
file_lo = 'tables/mzdové náklady bereko 2023.xlsx'

if all((Path(file_up).exists(), Path(file_lo).exists())):
    print('ok')
    inspector = Inspector(file_up, file_lo)
    print(inspector.refund_are_equal())
else:
    print('File does not exist')
