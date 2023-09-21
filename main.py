from sheet_transfer import SheetReaderWriter
from file_grabber import get_xl_files

def run(output_file):
    sheet_transfer = SheetReaderWriter(output_file, get_xl_files())
    sheet_transfer.cycle_through_sheets()


if __name__ == "__main__":
    run('result.xlsx')