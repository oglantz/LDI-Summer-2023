import tkinter as tk

class Window:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("LDI Budget Spreadsheet Tool")
        self.chosen_forecast_path = None
        self.chosen_forecast_path = None
        intro_text = tk.Label(self.window, text = "Welcome to the LDI Budged Spreadsheet Tool!")
        intro_text.grid(row = 0, column = 0, pady = 2, columnspan = 3)
        forecast_frame = tk.Frame(self.window)
        pipeline_frame = tk.Frame(self.window)
        empty_space = tk.Label(self.window, text = "             ")
        forecast_frame.grid(row = 1, column = 0, pady = 2)
        pipeline_frame.grid(row = 1, column = 2, pady = 2)
        empty_space.grid(row = 1, column = 1, pady = 2)
        select_pipeline_folder = tk.Button(pipeline_frame, text = "Select path for pipeline folder.")
        select_pipeline_folder.pack()
        select_forecast_folder = tk.Button(forecast_frame, text = "Select path for forecast folder.")
        select_forecast_folder.pack()
        self.chosen_pipeline_label = tk.Label(pipeline_frame, text = "Selected path will appear here.")
        self.chosen_pipeline_label.pack()
        self.chosen_forecast_label = tk.Label(forecast_frame, text = "Selected path will appear here.")
        self.chosen_forecast_label.pack()


    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    window = Window()
    window.run()

