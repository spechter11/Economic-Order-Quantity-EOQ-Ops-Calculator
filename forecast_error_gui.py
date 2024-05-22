import tkinter as tk
from tkinter import ttk, messagebox
from pandas import DataFrame
import os
from forecast_error_processor import ForecastErrorProcessor

class ForecastErrorApp:
    def __init__(self, parent):
        self.parent = parent
        self.create_error_widgets(parent)

    def create_error_widgets(self, parent):
        # Input fields for data
        ttk.Label(parent, text="Enter data for error calculation (Month-Year, Forecast, Demand) - space separated rows:").grid(column=0, row=0, columnspan=4, sticky=tk.W, padx=10, pady=5)
        self.error_data_entry = tk.Text(parent, height=10, width=80)
        self.error_data_entry.grid(column=0, row=1, columnspan=4, padx=10, pady=5)
        
        self.error_result = tk.Text(parent, height=10, width=80)
        self.error_result.grid(column=0, row=3, columnspan=4, padx=10, pady=5)
        
        ttk.Button(parent, text="Calculate Forecast Errors", command=self.calculate_errors).grid(column=0, row=2, columnspan=4, padx=10, pady=5)

        self.export_button = ttk.Button(parent, text="Export to Excel", command=self.export_errors)
        self.export_button.grid(column=0, row=4, columnspan=4, padx=10, pady=5)

    def calculate_errors(self):
        raw_data = self.error_data_entry.get("1.0", tk.END).strip()
        rows = raw_data.split('\n')
        data_dict = {'Month-Year (t)': [], 'Forecast (Ft)': [], 'Demand (Dt)': []}
        
        for row in rows:
            parts = row.split()
            month_year, forecast, demand = parts[0], parts[1], parts[2]
            data_dict['Month-Year (t)'].append(month_year.strip())
            data_dict['Forecast (Ft)'].append(float(forecast.strip().replace(',', '')))
            data_dict['Demand (Dt)'].append(float(demand.strip().replace(',', '')))

        self.df = pd.DataFrame(data_dict)
        self.calculator = ForecastErrorProcessor(self.df)
        results_df = self.calculator.generate_results_table()

        self.error_result.delete(1.0, tk.END)
        self.error_result.insert(tk.END, "Results Table:\n")
        self.error_result.insert(tk.END, self.df.to_string(index=False))
        self.error_result.insert(tk.END, "\n\nStatistics:\n")
        self.error_result.insert(tk.END, results_df.to_string(index=False))

    def export_errors(self):
        filename = os.path.join(os.path.expanduser("~"), "Downloads", "forecast_error_results.xlsx")
        self.calculator.export_to_excel(filename)
        messagebox.showinfo("Export Successful", f"Results exported to {filename}")
