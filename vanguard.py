import pandas as pd
from servant import clean, get_month, get_pension_type, get_ins_code
import openpyxl
import openpyxl.utils


class Vanguard:
    def __init__(self, file_mzdy, file_pracov):
        self.mzdy = file_mzdy
        self.pracov = file_pracov
        self.cats = []
        self.data = ''

    @property
    def prep_df(self):
        """ Preparation of data to dictionary. Uses pandas library to open, sort and merge data from csvs."""
        mzdy = pd.read_csv(self.mzdy, encoding='cp1250').applymap(clean)
        mzdy['Fare'] = mzdy[['Davky1', 'Davky2']].sum(axis=1, skipna=True)
        mzdy['Payout'] = mzdy[['Zamest', 'HrubaMzda', 'iNemoc']].sum(axis=1, skipna=True)
        mzdy['RokMes'] = mzdy['RokMes'].map(get_month)

        pracov = pd.read_csv(self.pracov, encoding='cp1250').applymap(clean)
        pracov['PensionType'] = pracov['TypDuch'].map(get_pension_type)

        data = pd.merge(mzdy, pracov, on='RodCislo', suffixes=('', '_y'))
        data = data.drop(['Kat_y', 'Kod_y', 'Jmeno30', 'Davky1', 'Davky2', 'Zamest', 'HrubaMzda', 'iNemoc'], axis=1)

        self.cats = list(set(data['Kat']))

        return data

    @property
    def merge_data(self):
        """Passes the dataframe to dictionary for easier parsing."""
        people = {}
        dataframe = self.prep_df

        # Convert dataframe to dictionary in form like this:
        # {idnum: {'name': str, ... ,'date: {list of dates with payments}}, idnum2: {...}, }
        for _, row in dataframe.iterrows():
            people.setdefault(row['RodCislo'], {
                'Name': row['JmenoS'],
                'Code': row['Kod'],
                'Cat': row['Kat'],
                'InsCode': get_ins_code(row['CisPoj']),
                'Date': {},
                'StartEmployment': row['VstupDoZam'],
                'EndEmployment': row['UkonceniZam'],
                'PensionType': row['PensionType'],
                'PensionStart': row['DuchOd']
            })
            people[row['RodCislo']]['Date'].setdefault(row['RokMes'], {
                'Fare': int(row['Fare']), 'Payout': int(row['Payout'])
            })
        return people


d = Vanguard(file_mzdy='data/Q2.CSV', file_pracov='data/PRACOVQ2.CSV')


class XScout:
    def __init__(self, spreadsheet):
        self.wb = openpyxl.load_workbook(spreadsheet)
        self.spread = self.spread()

    def employee_list(self):
        people = []
        for sheet_index, ws in enumerate(self.wb.worksheets[1:3]):
            sheet_ids = []
            for row in range(*self.spread[sheet_index]):
                sheet_ids.append(ws.cell(row, 4).value)
            people.append(sheet_ids)

        return people

    def first_row(self, sheet_number):
        ws = self.wb.worksheets[sheet_number]
        row = 1
        while True:
            if ws.cell(row, 1).value == 1:
                return row
            row += 1

    def last_row(self, sheet_number):
        ws = self.wb.worksheets[sheet_number]
        row = self.first_row(sheet_number)

        while True:
            if not ws.cell(row, 2).value:
                return row - 1
            row += 1

    def spread(self):
        s = []
        for sheet_number in range(1, 3):
            spread = (self.first_row(sheet_number), self.last_row(sheet_number))
            s.append(spread)
        return s


x = XScout('tables/jmenny_seznam_2022_09_27 Bereko.xlsx')
x.employee_list()
for sheet in x.employee_list():
    print('Sheet:')
    for i in sheet:
        print(i)


