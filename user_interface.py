import tkinter as tk
import tkinter.messagebox
from tkinter import filedialog
from pathlib import Path
import sheet_transfer

class Window:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("LDI Budget Spreadsheet Tool")
        self.chosen_dir_path = None
        intro_text = tk.Label(self.window, text = "Welcome to the LDI Budget Spreadsheet Tool!")
        intro_text.grid(row = 0, column = 0, pady = 2, columnspan = 3)
        directory_frame = tk.Frame(self.window)
        directory_frame.grid(row = 1, column = 1, pady = 2)
        select_pipeline_folder = tk.Button(directory_frame, text = "Select path for pipeline folder.", command = self._update_directory)
        select_pipeline_folder.pack()
        self.chosen_dir_label = tk.Label(directory_frame, text = "Selected path will appear here.")
        self.chosen_dir_label.pack()
        finish_button_instructions = tk.Label(self.window, text = "Insert desired file name into text box.")
        finish_button_instructions.grid(row = 2, column = 0, pady = 2, columnspan = 3)
        self.result_file_name_input = tk.Text(self.window, height = 1, width = 35)
        self.result_file_name_input.grid(row = 3, column = 0, pady = 2, columnspan = 3)
        finish_button = tk.Button(self.window, text = "Finish spreadsheet", command = self._create_sheet)
        finish_button.grid(row = 4, column = 0, pady = 2, columnspan = 3)
        warning_label = tk.Label(self.window, text = "WARNING: Do not press the finish button until you have selected the correct folder.", foreground = "red")
        warning_label.grid(row = 5, column = 0, pady = 2, columnspan = 3)

    def _update_directory(self):
        self.chosen_dir_path = filedialog.askdirectory()
        self.chosen_dir_label.configure(text = self.chosen_dir_path)
        self.chosen_dir_path = Path(self.chosen_dir_path)

    def _create_sheet(self):
        try:
            sheet_name = self.result_file_name_input.get(1.0, "end-1c")
            list_of_files = [f for f in self.chosen_dir_path.iterdir() if f.is_file and f.suffix == ".xlsx"]
            sheet_creator = sheet_transfer.SheetReaderWriter(sheet_name, list_of_files)
            sheet_creator.cycle_through_sheets()
            self.window.destroy()
        except Exception as e:
            tkinter.messagebox.showerror("Error", f"Unexpected error! Details below\n{type(e)}\n{e}")


    def run(self):
        self.window.mainloop()
