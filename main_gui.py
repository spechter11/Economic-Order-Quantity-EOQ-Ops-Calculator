import tkinter as tk
from tkinter import ttk, messagebox
from eop_gui import EOQApp
from forecast_gui import ForecastApp
from forecast_error_gui import ForecastErrorApp

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EOQ and Time Series Forecasting")
        self.configure_window()
        # Create a notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(padx=10, pady=10, fill='both', expand=True)

        # Create EOQ tab
        self.eoq_frame = ttk.Frame(self.notebook)
        EOQApp(self.eoq_frame, self.root, self.scale)
        self.notebook.add(self.eoq_frame, text='EOQ Calculator')

        # Create Time Series Forecast tab
        self.forecast_frame = ttk.Frame(self.notebook)
        ForecastApp(self.forecast_frame, self.scale)
        self.notebook.add(self.forecast_frame, text='Time Series Forecasting')

        # Create Forecast Error tab
        self.error_frame = ttk.Frame(self.notebook)
        ForecastErrorApp(self.error_frame)
        self.notebook.add(self.error_frame, text='Forecast Error Calculation')

    def configure_window(self):  # New method
        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Set window size relative to screen size
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        self.root.geometry(f"{window_width}x{window_height}")

        # Set scaling based on screen size
        self.scale = min(screen_width / 1024, screen_height / 800)
        self.root.tk.call('tk', 'scaling', self.scale)

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
