import tkinter as tk
from tkinter import ttk, messagebox
from eop_gui import EOQApp
from forecast_gui import ForecastApp
from forecast_error_gui import ForecastErrorApp

class MainApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EOQ and Time Series Forecasting")

        # Create a notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(padx=10, pady=10, fill='both', expand=True)

        # Create EOQ tab
        self.eoq_frame = ttk.Frame(self.notebook)
        EOQApp(self.eoq_frame)
        self.notebook.add(self.eoq_frame, text='EOQ Calculator')

        # Create Time Series Forecast tab
        self.forecast_frame = ttk.Frame(self.notebook)
        ForecastApp(self.forecast_frame)
        self.notebook.add(self.forecast_frame, text='Time Series Forecasting')

        # Create Forecast Error tab
        self.error_frame = ttk.Frame(self.notebook)
        ForecastErrorApp(self.error_frame)
        self.notebook.add(self.error_frame, text='Forecast Error Calculation')

if __name__ == "__main__":
    root = tk.Tk()
    app = MainApp(root)
    root.mainloop()
