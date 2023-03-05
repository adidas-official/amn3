import csv
import re
from io import BytesIO

import msoffcrypto
import openpyxl
from openpyxl.utils import get_column_letter

from copy import copy

mapping = [
    {
        0: 'J',
        1: 'O',
        2: 'T'
    },
    {
        0: 'I',
        1: 'J',
        2: 'K'
    }
]

column_map = [
    {
        0: 10,
        1: 15,
        2: 20
    },
    {
        0: 9,
        1: 10,
        2: 11
    }
]


def unlock(f, pwd):
    try:
        decrypted_wb = BytesIO()
        with open(f, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            office_file.load_key(password=pwd)
            office_file.decrypt(decrypted_wb)

        return openpyxl.load_workbook(filename=decrypted_wb)
    except (UnboundLocalError, msoffcrypto.exceptions.FileFormatError):
        return openpyxl.load_workbook(f)


def clean(text) -> str:
    if type(text) == str:
        return text.replace('\'', '').replace('/', '').strip()
    return text


def get_pension_type(pentype) -> str:
    if 'invalidní 1.' in pentype:
        return 'OZP12'
    elif 'invalidní 3.' in pentype:
        return 'TZP'
    else:
        return ''


def get_sheet_by_emp_data(workplace, emp_status):
    if emp_status == 'DPP':
        return 6
    elif emp_status == 'U':
        return 7

    workplaces = {
        'Sklad': 0,
        'Admin': 0,
        'Jedna': 0,
        'Dílna': 1,
        'Prode': 2,
        'ÚP So': 3,
        'KÚ Ch': 4,
        'KÚ': 5
    }

    return workplaces[workplace]


def get_month(month) -> str:
    m = month.split('.')
    return '.'.join((m[1], m[2]))


def get_ins_code(code):
    with open('insurance/Fiala_insurance_codes.csv', 'r') as f:
        codes = csv.reader(f)
        data = dict(codes)
        data = {clean(k): clean(v) for k, v in data.items()}
        if code in data:
            return data[code]


def from_df_to_dict(df, filt=False, pk='RodCislo'):
    """Passes the dataframe to dictionary for later parsing."""
    people = {}
    if filt:
        filtered = (~df['Kat'].str.contains('U'))
        df = df[filtered]

    # Convert dataframe to dictionary in form like this:
    # {idnum: {'name': str, ... ,'date: {list of dates with payments}}, idnum2: {...}, }
    for _, row in df.iterrows():
        people.setdefault(row[pk], {
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
        people[row[pk]]['Date'].setdefault(row['RokMes'], {
            'Fare': int(row['Fare']), 'Payout': int(row['Payout'])
        })
    return people


def insert_row(worksheet, row_index):
    # define range of cells to move
    max_col = get_column_letter(worksheet.max_column)
    max_row = worksheet.max_row

    table_name, table_range = worksheet.tables.items()[0]
    last_row = int(re.search(r'\d*$', table_range).group())
    new_range = re.sub(r'\d*$', str(int(last_row) + 1), table_range, count=1)

    # Move all cells from the row bellow the insertion point, including styles and translated formulas
    worksheet.move_range(f"A{str(row_index)}:{max_col}{max_row}", rows=1, cols=0, translate=True)

    # Define new range of table
    worksheet.tables[table_name].ref = new_range

    # Copy all cells from the row above the insertion point, including formatting and formulas
    for col in range(1, worksheet.max_column + 1):
        cell_above = worksheet.cell(row=row_index - 1, column=col)
        cell_to_copy = worksheet.cell(row=row_index, column=col)
        cell_to_copy.value = cell_above.value
        cell_to_copy.number_format = cell_above.number_format
        cell_to_copy.font = copy(cell_above.font)
        if cell_above.has_style:
            cell_to_copy.border = copy(cell_above.border)
            cell_to_copy.fill = copy(cell_above.fill)
            cell_to_copy.number_format = copy(cell_above.number_format)
            cell_to_copy.protection = copy(cell_above.protection)
            cell_to_copy.alignment = copy(cell_above.alignment)
