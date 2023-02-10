import csv


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

