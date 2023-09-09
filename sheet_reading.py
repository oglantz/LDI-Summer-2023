import openpyxl
from pycel.excelcompiler import ExcelCompiler

class SheetReader:
    def __init__(self, spreadsheet_filename):
        self._path = spreadsheet_filename
        self._book = openpyxl.load_workbook(self._path)
        self._sheet = self._book.sheetnames[0]

    def get_cell_value(self, address: str):
        if self.cell_has_formula(address):
            return self.evaluate(address)
        else:
            return self.return_cell(address)

    def cell_has_formula(self, address: str):
        cell = self._book[self._sheet][address]
        if type(cell.value) == str and cell.value[0] == "=":
            return True
        else:
            return False

    def evaluate(self, address):
        compiler = ExcelCompiler(self._path)
        return compiler.evaluate(f"{self._sheet}!{address}")

    def return_cell(self, address):
        return self._book[self._sheet][address].value

    def collect_range_of_cells(self, start_cell, end_cell):
        return self._book[self._sheet][start_cell: end_cell]

    def get_all_cell_values(self, cell_range):
        final = []
        temp_row = []
        for row in cell_range:
            for cell in row:
                if self.cell_has_formula(cell.coordinate):
                    temp_row.append(self.evaluate(cell.coordinate))
                else:
                    temp_row.append(cell.value)
            final.append(temp_row)
            temp_row = []
        return final



# if __name__ == "__main__":
#     sheet_reader = SheetReader('IT LEAP Budget FY23-24 SPRING.xlsx')
#     # print(sheet_reader.get_cell_value("K4"))
#     cell_range = sheet_reader.collect_range_of_cells('A1', 'L4')
#     print(sheet_reader.get_all_cell_values(cell_range))