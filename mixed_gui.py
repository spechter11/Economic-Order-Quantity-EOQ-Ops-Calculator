import tkinter as tk
from tkinter.ttk import Notebook, Frame, Label, Combobox, Button
from tkinter import messagebox
from os import path, expanduser
from pandas import DataFrame
from eop_processor import EOQProcessor
from time_series_calculations import TimeSeriesForecast
from forecast_error_processor import ForecastErrorProcessor


class EOQApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EOQ and Time Series Forecasting")

        # Create a notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(padx=10, pady=10, fill='both', expand=True)

        # Create EOQ tab
        self.eoq_frame = ttk.Frame(self.notebook)
        self.create_eoq_widgets(self.eoq_frame)
        self.notebook.add(self.eoq_frame, text='EOQ Calculator')

        # Create Time Series Forecast tab
        self.forecast_frame = ttk.Frame(self.notebook)
        self.create_forecast_widgets(self.forecast_frame)
        self.notebook.add(self.forecast_frame, text='Time Series Forecasting')

        # Create Forecast Error tab
        self.error_frame = ttk.Frame(self.notebook)
        self.create_error_widgets(self.error_frame)
        self.notebook.add(self.error_frame, text='Forecast Error Calculation')

    def create_eoq_widgets(self, parent):
        # Labels and dropdown menus for inputs
        self.entries = {}
        parameters = [
            ("Demand Rate (units/week)", "demand_rate"),
            ("Demand Yearly (units/year)", "demand_yearly"),
            ("Purchase Cost (dollars/unit)", "purchase_cost"),
            ("Holding Cost Rate (annual %)", "holding_cost_rate"),
            ("Ordering Cost (dollars/order)", "ordering_cost"),
            ("Weeks per Year", "weeks_per_year"),
            ("EOQ", "EOQ")
        ]

        for idx, (label_text, var_name) in enumerate(parameters):
            label = ttk.Label(parent, text=label_text)
            label.grid(row=idx, column=0, padx=10, pady=5, sticky=tk.W)
            
            var = tk.StringVar(value="None")
            dropdown = ttk.Combobox(parent, textvariable=var, values=["None"])
            dropdown.grid(row=idx, column=1, padx=10, pady=5, sticky=tk.W)
            self.entries[var_name] = dropdown

        # Buttons
        self.calculate_button = ttk.Button(parent, text="Calculate", command=self.calculate)
        self.calculate_button.grid(row=len(parameters), column=0, padx=10, pady=10, sticky=tk.W)

        self.plot_button = ttk.Button(parent, text="Visualize", command=self.visualize)
        self.plot_button.grid(row=len(parameters), column=1, padx=10, pady=10, sticky=tk.W)

        # Text area for displaying results
        self.results_text = tk.Text(parent, height=10, width=80)
        self.results_text.grid(row=len(parameters) + 1, column=0, columnspan=3, padx=10, pady=10, sticky=tk.W)
        
        # Path for saved Excel file
        self.excel_path = os.path.join(os.path.expanduser("~"), "Downloads", "eoq_results.xlsx")

    def get_input_values(self):
        inputs = {}
        for var_name, dropdown in self.entries.items():
            value = dropdown.get()
            if value == "None":
                inputs[var_name] = None
            else:
                try:
                    inputs[var_name] = float(value) if var_name not in ["weeks_per_year", "EOQ"] else int(value)
                except ValueError:
                    inputs[var_name] = None
        return inputs

    def calculate(self):
        inputs = self.get_input_values()
        processor = EOQProcessor(**inputs)
        input_table = processor.generate_input_table()
        results_table = processor.generate_results_table()
        self.display_results(input_table, results_table)
        processor.export_to_excel(self.excel_path)
        messagebox.showinfo("Download Successful", "Downloaded Solution Excel. Check your Downloads folder.")

    def display_results(self, input_table, results_table):
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Inputs Table:\n")
        self.results_text.insert(tk.END, input_table.to_string(index=False))
        self.results_text.insert(tk.END, "\n\nResults Table:\n")
        self.results_text.insert(tk.END, results_table.to_string(index=False))

    def visualize(self):
        inputs = self.get_input_values()
        processor = EOQProcessor(**inputs)
        processor.plot_costs()

    def create_forecast_widgets(self, parent):
        # Input fields for data
        ttk.Label(parent, text="Enter time series data (comma separated values):").grid(column=0, row=0, sticky=tk.W, padx=10, pady=5)
        self.data_entry = tk.Entry(parent, width=50)
        self.data_entry.grid(column=1, row=0, padx=10, pady=5)

        # Simple Moving Average
        ttk.Label(parent, text="Simple Moving Average: Enter the window size (number of periods):").grid(column=0, row=1, sticky=tk.W, padx=10, pady=5)
        self.sma_window_entry = tk.Entry(parent, width=10)
        self.sma_window_entry.grid(column=1, row=1, padx=10, pady=5)
        self.sma_result = ttk.Label(parent, text="Result will appear here")
        self.sma_result.grid(column=2, row=1, padx=10, pady=5)
        ttk.Button(parent, text="Calculate Simple Moving Average", command=self.calculate_sma).grid(column=3, row=1, padx=10, pady=5)

        # Weighted Moving Average
        ttk.Label(parent, text="Weighted Moving Average: Enter weights (comma separated, sum should be 1):").grid(column=0, row=2, sticky=tk.W, padx=10, pady=5)
        self.wma_weights_entry = tk.Entry(parent, width=50)
        self.wma_weights_entry.grid(column=1, row=2, padx=10, pady=5)
        self.wma_result = ttk.Label(parent, text="Result will appear here")
        self.wma_result.grid(column=2, row=2, padx=10, pady=5)
        ttk.Button(parent, text="Calculate Weighted Moving Average", command=self.calculate_wma).grid(column=3, row=2, padx=10, pady=5)

        # Exponential Smoothing
        ttk.Label(parent, text="Exponential Smoothing: Enter the smoothing factor (alpha):").grid(column=0, row=3, sticky=tk.W, padx=10, pady=5)
        self.es_alpha_entry = tk.Entry(parent, width=10)
        self.es_alpha_entry.grid(column=1, row=3, padx=10, pady=5)
        ttk.Label(parent, text="Enter the prior forecast value:").grid(column=0, row=4, sticky=tk.W, padx=10, pady=5)
        self.es_prior_entry = tk.Entry(parent, width=10)
        self.es_prior_entry.grid(column=1, row=4, padx=10, pady=5)
        ttk.Label(parent, text="Enter the observed demand for the last period:").grid(column=0, row=5, sticky=tk.W, padx=10, pady=5)
        self.es_observed_entry = tk.Entry(parent, width=10)
        self.es_observed_entry.grid(column=1, row=5, padx=10, pady=5)
        self.es_result = ttk.Label(parent, text="Result will appear here")
        self.es_result.grid(column=2, row=5, padx=10, pady=5)
        ttk.Button(parent, text="Calculate Exponential Smoothing", command=self.calculate_es).grid(column=3, row=5, padx=10, pady=5)

    def calculate_sma(self):
        data = self.get_data()
        sma_window = int(self.sma_window_entry.get())
        ts_forecast = TimeSeriesForecast(data)
        result = ts_forecast.simple_moving_average(sma_window)
        self.sma_result.config(text=f"Simple Moving Average: {result:.2f}")

    def calculate_wma(self):
        data = self.get_data()
        weights_str = self.wma_weights_entry.get()
        weights = [float(w) for w in weights_str.split(",")]
        ts_forecast = TimeSeriesForecast(data)
        result = ts_forecast.weighted_moving_average(weights)
        self.wma_result.config(text=f"Weighted Moving Average: {result:.2f}")

    def calculate_es(self):
        data = self.get_data()
        alpha = float(self.es_alpha_entry.get())
        prior_forecast = float(self.es_prior_entry.get())
        observed_demand = float(self.es_observed_entry.get())
        ts_forecast = TimeSeriesForecast(data)
        result = ts_forecast.exponential_smoothing(alpha, prior_forecast, observed_demand)
        self.es_result.config(text=f"Exponential Smoothing: {result:.2f}")

    def get_data(self):
        data_str = self.data_entry.get()
        # Remove commas and convert to list of floats
        data = [float(i.replace(',', '')) for i in data_str.split(",")]
        return data

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

if __name__ == "__main__":
    root = tk.Tk()
    app = EOQApp(root)
    root.mainloop()