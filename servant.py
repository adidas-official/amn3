import csv
from io import BytesIO

import msoffcrypto
import openpyxl

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
