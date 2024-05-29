import tkinter as tk
from tkinter import ttk, messagebox
from pandas import DataFrame
import pandas as pd
import os
from forecast_error_processor import ForecastErrorProcessor

class ForecastErrorApp:
    def __init__(self, parent):
        self.parent = parent
        self.create_error_widgets(parent)

    def create_error_widgets(self, parent):
        # Instructions label
        instructions = (
            "Enter data for error calculation (Month-Year, Forecast, Demand) - space separated rows:"
            "\nExample:"
            "\nJan-2023 1000 1100"
            "\nFeb-2023 1200 1150"
            "\nMar-2023 1300 1250"
        )
        ttk.Label(parent, text=instructions, justify=tk.LEFT, font=("Helvetica", 12, "bold")).grid(column=0, row=0, columnspan=4, sticky=tk.W, padx=10, pady=10)

        # Input fields for data
        self.error_data_entry = tk.Text(parent, height=10, width=80, font=("Helvetica", 12))
        self.error_data_entry.grid(column=0, row=1, columnspan=4, padx=10, pady=5)
        self.error_data_entry.insert(tk.END, "Jan-2023 1000 1100\nFeb-2023 1200 1150\nMar-2023 1300 1250")  # Sample data

        # Calculate button
        ttk.Button(parent, text="Calculate Forecast Errors", command=self.calculate_errors).grid(column=0, row=2, columnspan=4, padx=10, pady=10)

        # Result display
        self.error_result = tk.Text(parent, height=10, width=80, font=("Helvetica", 12), state='disabled')
        self.error_result.grid(column=0, row=3, columnspan=4, padx=10, pady=5)

        # Export button
        self.export_button = ttk.Button(parent, text="Export to Excel", command=self.export_errors, state='disabled')
        self.export_button.grid(column=0, row=4, columnspan=4, padx=10, pady=10)

    def calculate_errors(self):
        raw_data = self.error_data_entry.get("1.0", tk.END).strip()
        rows = raw_data.split('\n')
        data_dict = {'Month-Year (t)': [], 'Forecast (Ft)': [], 'Demand (Dt)': []}
        
        try:
            for row in rows:
                parts = row.split()
                month_year, forecast, demand = parts[0], parts[1], parts[2]
                data_dict['Month-Year (t)'].append(month_year.strip())
                data_dict['Forecast (Ft)'].append(float(forecast.strip().replace(',', '')))
                data_dict['Demand (Dt)'].append(float(demand.strip().replace(',', '')))
            
            self.df = pd.DataFrame(data_dict)
            self.calculator = ForecastErrorProcessor(self.df)
            results_df = self.calculator.generate_results_table()
            
            self.error_result.configure(state='normal')
            self.error_result.delete(1.0, tk.END)
            self.error_result.insert(tk.END, "Results Table:\n")
            self.error_result.insert(tk.END, self.df.to_string(index=False))
            self.error_result.insert(tk.END, "\n\nStatistics:\n")
            self.error_result.insert(tk.END, results_df.to_string(index=False))
            self.error_result.configure(state='disabled')

            self.export_button.configure(state='normal')
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def export_errors(self):
        try:
            filename = os.path.join(os.path.expanduser("~"), "Downloads", "forecast_error_results.xlsx")
            self.calculator.export_to_excel(filename)
            messagebox.showinfo("Export Successful", f"Results exported to {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during export: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Forecast Error Calculator")
    app = ForecastErrorApp(root)
    root.mainloop()
