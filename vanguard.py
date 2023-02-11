import pandas as pd
from servant import clean, get_month, get_pension_type, get_ins_code
import openpyxl
import openpyxl.utils


class Vanguard:
    def __init__(self, file_mzdy, file_pracov):
        self.mzdy = file_mzdy
        self.pracov = file_pracov

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


class XScout:
    def __init__(self, spreadsheet):
        self.wb = openpyxl.load_workbook(spreadsheet)
        self.spread = self.spread()

    def employee_list(self):
        people = []
        for sheet_index, ws in enumerate(self.wb.worksheets[1:3]):
            sheet_ids = []
            for row in range(*self.spread[sheet_index]):
                sheet_ids.append(clean(ws.cell(row, 4).value))
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


x = XScout('tables/jmenny_seznam_2022_09_27 Fiala.xlsx')
d = Vanguard(file_mzdy='data/Q2.CSV', file_pracov='data/PRACOVQ2.CSV')

# Dictionary with employees data from exported CSVs merged into single object.
"""
{
    idnum1: {
        name: 'abc',
        date: {
            month: {
                fare: 1234,
                payout: 9886
            }
        }
        other_data: ...
    },
    idnum2: ...,
}
"""
merged = d.merge_data

# Converted to set for comparison of keys
merged_set = set(merged)

# Employees present on spreadsheet for this term.
# Distrubuted to two lists. First is `jmenny_seznam`, second is `nakl.prov`
employee_list = x.employee_list()

# List of all employees from both sheets in xlsx.
list_of_all_employees_in_xlsx = set([i for sublist in employee_list for i in sublist])
# People who are NOT in exported CSVs, but ARE in xlsx. They don't have a payout for this term
xlsx_minus_export = list_of_all_employees_in_xlsx.difference(merged_set)
[print(i) for i in xlsx_minus_export]

print()
# People who are NOT in xlsx. New people who are not yet present on the list of employee for this term.
export_minus_xlsx = merged_set.difference(list_of_all_employees_in_xlsx)
[print(i) for i in export_minus_xlsx]
