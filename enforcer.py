from itertools import chain

from vanguard import Vanguard
from servant import mapping
from openpyxl.utils import get_column_letter


def to_list(idnum, items, row_num, num):

    message = f'{idnum} {items["Name"]} is in sheet{num} row #{row_num}.'
    for month, payout_data in items["Date"].items():
        m = (int(month.split('.')[0]) + 2) % 3

        month_col = mapping[num-2][m]

        message += f'\n-Payout for month {month}: {payout_data["Payout"]}=>{month_col}{row_num}'

    print(message)


def new_person(idnum, items):
    message = f'{idnum} {items["Name"]} is new and will be placed to list'
    pen_t = items["PensionType"]
    if pen_t:
        message += ' 2'
    else:
        message += ' 3'


class Enforcer:
    def __init__(self, dataset):
        self.data_up = dataset[0][0]
        self.data_lo = dataset[0][1]
        self.merged_up = dataset[1][0]
        self.merged_lo = dataset[1][1]
        self.x = dataset[-1]
        self.last_row = [row[1] + 1 for row in self.x.range]

    def display_data(self):
        for person, data in self.merged_up.items():
            if person in self.data_up[0]:
                to_list(person, data, self.data_up[0][person], 2)
            elif person in self.data_up[1]:
                to_list(person, data, self.data_up[1][person], 3)
            else:
                print(f'{person} {data["Name"]} is new')
                if data['PensionType'] != '':
                    to_list(person, data, self.last_row[0], 2)
                    self.last_row[0] += 1
                    # message += 'belongs to list2'
                else:
                    to_list(person, data, self.last_row[1], 3)
                    self.last_row[1] += 1
                    # message += 'belongs to list3'
                # print(message)

    def display_lo(self):
        for person, data in self.merged_lo.items():
            person_found = False
            for i, sheet_data in enumerate(self.data_lo):
                if i < 3:
                    fare_shift = 4
                else:
                    fare_shift = 2

                if person in sheet_data:
                    person_found = True
                    # print(i, data, sheet_data[person])
                    message = f'Sheet: {i}, Line: {sheet_data[person]}, Person: {person}'
                    for date, money in data['Date'].items():
                        col = self.x.get_month(date, i)
                        fare_col = col + fare_shift
                        col_letter = get_column_letter(col)
                        fare_letter = get_column_letter(fare_col)
                        message += f'\n- {col_letter}{sheet_data[person]}: {money["Payout"]}'
                        if money["Fare"]:
                            message += f', {fare_letter}{sheet_data[person]}: {money["Fare"]}'
                    print(message)
                # else:
                #     print(f'Person {person} not in {i}')
            if not person_found:
                print(f'Person {person} is new and should be written to {data["Code"]} sheet')

    def get_missing(self):
        # merged_lo is dictionary with all data for each employee
        # 'Arvensisová Radka': {'Name': 'Arvensisová Radka', 'Code': 'Prode',...
        # data_lo is list of dictionaries with key:value pairs being 'name':'line_number' for each sheet
        # [{'Bobok Vilém': 3, 'Cenefels Jan': 4, 'Dbalý Petr': 5, 'Diviak Miroslav': 6, ...
        data_for_month = set(self.merged_lo.keys())
        names_in_xlsx = set(chain.from_iterable(d.keys() for d in self.data_lo))
        print('People without pay for this month')
        print(names_in_xlsx.difference(data_for_month))
        print(self.get_last_rows())
        print('New people')
        for name in data_for_month.difference(names_in_xlsx):
            print(self.merged_lo[name]['Name'], self.merged_lo[name]['Code'])
            for month, money in self.merged_lo[name]['Date'].items():
                print(month.split('.')[0], money['Payout'], money['Fare'])

    def show_all(self):
        for name, data in self.merged_lo.items():
            print(name, data)

    def get_last_rows(self):
        last_rows = [self.x.last_row_lo(i) for i in range(8)]
        return last_rows


vanguard = Vanguard(file_mzdy='data/Q2.CSV', file_pracov='data/PRACOVQ2.CSV')
enforcer = Enforcer(vanguard.loader)
# enforcer.display_data()
# enforcer.display_lo()
enforcer.get_missing()
