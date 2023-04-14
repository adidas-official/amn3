import pandas as pd
from . import servant
import openpyxl
import openpyxl.utils
from . import paths
from pathlib import Path


class Assembler:
    def __init__(self, file_mzdy, file_pracov):
        self.mzdy = file_mzdy
        self.pracov = file_pracov
        self.dataframe = self.prep_df

    @property
    def prep_df(self):
        """ Preparation of data to dictionary. Uses pandas library to open, sort and merge data from csvs."""
        mzdy = pd.read_csv(self.mzdy, encoding='cp1250').applymap(servant.clean)
        mzdy['Fare'] = mzdy[['Davky1', 'Davky2']].sum(axis=1, skipna=True)
        mzdy['Payout'] = mzdy[['Zamest', 'HrubaMzda', 'iNemoc']].sum(axis=1, skipna=True)
        mzdy['RokMes'] = mzdy['RokMes'].map(servant.get_month)

        pracov = pd.read_csv(self.pracov, encoding='cp1250').applymap(servant.clean)
        pracov['PensionType'] = pracov['TypDuch'].map(servant.get_pension_type)

        data = pd.merge(mzdy, pracov, on='RodCislo', suffixes=('', '_y'))
        data = data.drop(['Kat_y', 'Kod_y', 'Jmeno30', 'Davky1', 'Davky2', 'Zamest', 'HrubaMzda', 'iNemoc'], axis=1)

        return data

    @property
    def loader(self):
        """ Running all important pieces together. Creating list of employees in spreadsheets. """

        # Dictionary with employees data from exported CSVs merged into single object.
        # { idnum1: { name: 'abc', date: { month: { fare: 1234, payout: 9886 } } other_data: ... }, idnum2: ..., }
        merged_lists = (servant.from_df_to_dict(self.dataframe, True, 'RodCislo'),
                        servant.from_df_to_dict(self.dataframe, False, 'JmenoS'))

        scout = Scout(Path(paths.TABLES_PATH) / 'jmenny_seznam_2022_09_27 Fiala.xlsx', Path(paths.TABLES_PATH) / 'Mzdové náklady 2023.xlsx')

        employee_lists = (scout.employee_list_up(), scout.employee_list_lo())

        return employee_lists, merged_lists, scout, servant.get_q(self.dataframe)


class Scout:
    def __init__(self, spreadsheet1, spreadsheet2):
        self.wb_up = openpyxl.load_workbook(spreadsheet1)
        self.wb_lo = servant.unlock(spreadsheet2, '13881744')
        self.range = self.spread

    def employee_list_up(self):
        """ Returns list of people present on spreadsheet. Each sheet has its own dictionary with person:row kw pair"""
        people = []
        for sheet_index, ws in enumerate(self.wb_up.worksheets[1:3]):
            sheet_ids = {}
            for row in range(self.range[sheet_index][0], self.range[sheet_index][1] + 1):
                data_row = servant.clean(ws.cell(row, 4).value)
                sheet_ids[data_row] = row
            people.append(sheet_ids)

        return people

    def first_row(self, sheet_number):
        ws = self.wb_up.worksheets[sheet_number]
        row = 1
        while True:
            if ws.cell(row, 1).value == 1:
                return row
            row += 1

    def last_row(self, sheet_number):
        ws = self.wb_up.worksheets[sheet_number]
        row = self.first_row(sheet_number)

        while True:
            if not ws.cell(row, 2).value:
                return row - 1
            row += 1

    def last_row_lo(self, sheet_number):
        ws = self.wb_lo.worksheets[sheet_number]
        row = 3

        while True:
            if not ws.cell(row, 2).value:
                return row - 1
            row += 1

    @property
    def spread(self):
        s = []
        for sheet_number in range(1, 3):
            spread = (self.first_row(sheet_number), self.last_row(sheet_number))
            s.append(spread)
        return s

    def employee_list_lo(self):
        """ Returns list of people present on spreadsheet. Each sheet has its own dictionary with person:row kw pair"""
        people = []
        for sheet_index, ws in enumerate(self.wb_lo.worksheets[:-2]):
            row = 3
            sheet_names = {}
            while True:
                name = str(ws.cell(row, 2).value)[:20]
                if name == '[ENDBLOCK]':
                    break

                if name:
                    sheet_names[name] = row
                row += 1

            people.append(sheet_names)

        return people

    def get_month(self, date, sheet_num):
        """ Gets month column index in local table from date in employee data object """
        # Counter for keeping track of curent column index
        counter = 0
        # Number of how many times merged cell was found
        m = 0
        month_num = int(date.split('.')[0])
        cell = False

        while True:
            if m == month_num:
                return cell.column

            cell = self.wb_lo.worksheets[sheet_num].cell(row=1, column=counter + 3)
            counter += 1

            if not type(cell).__name__ == 'MergedCell' and cell.value:
                m += 1
