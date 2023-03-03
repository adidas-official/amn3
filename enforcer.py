from vanguard import Vanguard
from servant import mapping


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
        self.last_row = [row[1] + 1 for row in dataset[-1]]

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
            for i, sheet_data in enumerate(self.data_lo):
                if person in sheet_data:
                    print(i, data)


vanguard = Vanguard(file_mzdy='data/Q2.CSV', file_pracov='data/PRACOVQ2.CSV')
enforcer = Enforcer(vanguard.loader)
# enforcer.display_data()
enforcer.display_lo()

