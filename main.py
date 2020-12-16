from Columns import columns
from openpyxl import load_workbook

"""
    Function will added columns to the excel document
"""


class Workbook_Worker():
    def __init__(self, workbook):
        self.wb = load_workbook(f"{workbook}.xlsx")
        self.sheet = self.wb.active
        self.columns_in_sheet = []
        # self.add_column()
        self.rows = columns[0].split('\n')
        """ Save the workbook """
        self.read_columns()
        self.compare_columns()
        # self.wb.save(f'{workbook}_new.xlsx')

    def add_column(self):
        i = 1
        for x in self.rows:
            self.sheet.cell(column=2, row=i).value = x
            i += 1

    def read_columns(self):
        """
            We iterate through all the columns append A1,B1,... so we read every colummn
            we then append all the values of each column into another 2D array
            We iterate through the max row in the excel sheet, if a cell has no value
            it is None
        """
        letter = 'A'
        for index in range(0, len(columns)+1):
            self.columns_in_sheet.append([])
            for row in range(1, self.sheet.max_row+1):
                self.columns_in_sheet[index].append(
                    self.sheet[f"{letter}{row}"].value)
            letter = chr(ord(letter) + 1)

    def compare_columns(self):
        unique_values = []
        i = 0
        for col in self.columns_in_sheet:
            unique_values.append([])
            for value in col:
                for second_column in self.columns_in_sheet:
                    found = False
                    if(value in second_column and second_column != col or value == None):
                        found = True
                        break
                if(not found):
                    unique_values[i].append(value)
            i += 1
        print(unique_values)


if __name__ == '__main__':
    Workbook_Worker('numbers')