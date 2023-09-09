from sheet_reading import SheetReader
from openpyxl import Workbook
from file_grabber import get_xl_files

class SheetWriter:

    def __init__(self, result_file: str, file_list: list):
        self._final_file = result_file
        self._file_list = file_list
        self._empty_row = [None] * 12

        self._final_book = Workbook()
        self._final_sheet = self._final_book.active

    def save_file(self):
        self._final_book.save(self._final_file)
        print(f'File successfully saved as {self._final_file}')

    def write_sheet_data(self):
        for file in self._file_list:
            sheet_reader = SheetReader(file)
            cell_range = sheet_reader.collect_range_of_cells('A1', 'L4')
            temp_data = sheet_reader.get_all_cell_values(cell_range)
            for row in temp_data:
                self._final_sheet.append(row)
            self._final_sheet.append(self._empty_row)
            self._final_sheet.append(self._empty_row)

    def run(self):
        self.write_sheet_data()
        self.save_file()


if __name__ == "__main__":
    writer = SheetWriter('result.xlsx', get_xl_files())
    writer.run()
