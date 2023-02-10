import pandas as pd
from servant import clean, get_month, get_pension_type, get_ins_code


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
print(d.merge_data)
