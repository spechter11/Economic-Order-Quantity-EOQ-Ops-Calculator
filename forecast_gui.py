import tkinter as tk
from tkinter import ttk, messagebox
from pandas import DataFrame, ExcelWriter
from numpy import array
import os
from time_series_forecast import TimeSeriesForecast


class ForecastApp:
    def __init__(self, parent):
        self.parent = parent
        self.sma_result_value = None
        self.wma_result_value = None
        self.es_result_value = None
        self.create_forecast_widgets(parent)

    def create_forecast_widgets(self, parent):
        # Input fields for data
        ttk.Label(parent, text="Enter time series data (space separated rows):").grid(column=0, row=0, sticky=tk.W, padx=10, pady=5)
        self.data_entry = tk.Text(parent, height=10, width=50)
        self.data_entry.grid(column=1, row=0, padx=10, pady=5)

        # Simple Moving Average
        ttk.Label(parent, text="Simple Moving Average: Enter the window size (number of periods):").grid(column=0, row=1, sticky=tk.W, padx=10, pady=5)
        self.sma_window_entry = tk.Entry(parent, width=10)
        self.sma_window_entry.grid(column=1, row=1, padx=10, pady=5)
        self.sma_result = ttk.Label(parent, text="Result will appear here")
        self.sma_result.grid(column=2, row=1, padx=10, pady=5)
        ttk.Button(parent, text="Calculate Simple Moving Average", command=self.calculate_sma).grid(column=3, row=1, padx=10, pady=5)
        ttk.Button(parent, text="Export SMA to Excel", command=self.export_sma_to_excel).grid(column=4, row=1, padx=10, pady=5)

        # Weighted Moving Average
        ttk.Label(parent, text="Weighted Moving Average: Enter weights (space separated, sum should be 1):").grid(column=0, row=2, sticky=tk.W, padx=10, pady=5)
        self.wma_weights_entry = tk.Text(parent, height=10, width=50)
        self.wma_weights_entry.grid(column=1, row=2, padx=10, pady=5)
        self.wma_result = ttk.Label(parent, text="Result will appear here")
        self.wma_result.grid(column=2, row=2, padx=10, pady=5)
        ttk.Button(parent, text="Calculate Weighted Moving Average", command=self.calculate_wma).grid(column=3, row=2, padx=10, pady=5)
        ttk.Button(parent, text="Export WMA to Excel", command=self.export_wma_to_excel).grid(column=4, row=2, padx=10, pady=5)

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
        ttk.Button(parent, text="Export ES to Excel", command=self.export_es_to_excel).grid(column=4, row=5, padx=10, pady=5)

        # Export All
        ttk.Button(parent, text="Export All to Excel", command=self.export_all_to_excel).grid(column=0, row=6, columnspan=5, padx=10, pady=20)

    def calculate_sma(self):
        data = self.get_data()
        sma_window = int(self.sma_window_entry.get())
        ts_forecast = TimeSeriesForecast(data)
        result = ts_forecast.simple_moving_average(sma_window)
        self.sma_result.config(text=f"Simple Moving Average: {result:.2f}")
        self.sma_result_value = result

    def calculate_wma(self):
        data = self.get_data()
        weights = self.get_weights()
        ts_forecast = TimeSeriesForecast(data)
        result = ts_forecast.weighted_moving_average(weights)
        self.wma_result.config(text=f"Weighted Moving Average: {result:.2f}")
        self.wma_result_value = result

    def calculate_es(self):
        data = self.get_data()
        alpha = self.parse_number(self.es_alpha_entry.get())
        prior_forecast = self.parse_number(self.es_prior_entry.get())
        observed_demand = self.parse_number(self.es_observed_entry.get())
        ts_forecast = TimeSeriesForecast(data)
        result = ts_forecast.exponential_smoothing(alpha, prior_forecast, observed_demand)
        self.es_result.config(text=f"Exponential Smoothing: {result:.2f}")
        self.es_result_value = result

    def get_data(self):
        raw_data = self.data_entry.get("1.0", tk.END).strip()
        data = [float(row.split()[1].replace(',', '')) for row in raw_data.split('\n')]
        return np.array(data)

    def get_weights(self):
        raw_weights = self.wma_weights_entry.get("1.0", tk.END).strip()
        weights = [float(row.split()[1]) for row in raw_weights.split('\n')]
        return np.array(weights)

    def parse_number(self, num_str):
        return float(num_str.replace(',', ''))

    def export_sma_to_excel(self):
        data = self.get_data()
        dates = [row.split()[0] for row in self.data_entry.get("1.0", tk.END).strip().split('\n')]
        df = pd.DataFrame({'Date': dates, 'Demand (Dt)': data})
        forecast_df = pd.DataFrame({
            'Date': ['Next Period'],
            'Simple Moving Average Forecast': [self.sma_result_value]
        })
        self.export_to_excel('Simple Moving Average', df, forecast_df, None, 'sma_result.xlsx')

    def export_wma_to_excel(self):
        data = self.get_data()
        weights = self.get_weights()
        dates = [row.split()[0] for row in self.data_entry.get("1.0", tk.END).strip().split('\n')]
        df = pd.DataFrame({'Date': dates, 'Demand (Dt)': data})
        weights_labels, weights_values = zip(*[row.split() for row in self.wma_weights_entry.get("1.0", tk.END).strip().split('\n')])
        weights_df = pd.DataFrame({'Weight #': weights_labels, 'Weight': [float(value) for value in weights_values]})
        forecast_df = pd.DataFrame({
            'Date': ['Next Period'],
            'Weighted Moving Average Forecast': [self.wma_result_value]
        })
        self.export_to_excel('Weighted Moving Average', df, weights_df, forecast_df, 'wma_result.xlsx')

    def export_es_to_excel(self):
        alpha = self.parse_number(self.es_alpha_entry.get())
        prior_forecast = self.parse_number(self.es_prior_entry.get())
        observed_demand = self.parse_number(self.es_observed_entry.get())
        es_df = pd.DataFrame({
            'Parameter': ['Smoothing Factor (alpha)', 'Prior Forecast', 'Observed Demand', 'Next Period Forecast'],
            'Value': [alpha, prior_forecast, observed_demand, self.es_result_value]
        })
        self.export_to_excel('Exponential Smoothing', es_df, None, None, 'es_result.xlsx')

    def export_all_to_excel(self):
        data = self.get_data()
        dates = [row.split()[0] for row in self.data_entry.get("1.0", tk.END).strip().split('\n')]

        # Calculate SMA if not already calculated
        if self.sma_result_value is None:
            sma_window = int(self.sma_window_entry.get())
            ts_forecast = TimeSeriesForecast(data)
            self.sma_result_value = ts_forecast.simple_moving_average(sma_window)

        # Simple Moving Average
        sma_df = pd.DataFrame({'Date': dates, 'Demand (Dt)': data})
        sma_forecast_df = pd.DataFrame({
            'Date': ['Next Period'],
            'Simple Moving Average Forecast': [self.sma_result_value]
        })

        # Calculate WMA if not already calculated
        if self.wma_result_value is None:
            weights = self.get_weights()
            ts_forecast = TimeSeriesForecast(data)
            self.wma_result_value = ts_forecast.weighted_moving_average(weights)

        # Weighted Moving Average
        weights = self.get_weights()
        weights_labels, weights_values = zip(*[row.split() for row in self.wma_weights_entry.get("1.0", tk.END).strip().split('\n')])
        wma_df = pd.DataFrame({'Date': dates, 'Demand (Dt)': data})
        weights_df = pd.DataFrame({'Weight #': weights_labels, 'Weight': [float(value) for value in weights_values]})
        wma_forecast_df = pd.DataFrame({
            'Date': ['Next Period'],
            'Weighted Moving Average Forecast': [self.wma_result_value]
        })

        # Calculate ES if not already calculated
        if self.es_result_value is None:
            alpha = self.parse_number(self.es_alpha_entry.get())
            prior_forecast = self.parse_number(self.es_prior_entry.get())
            observed_demand = self.parse_number(self.es_observed_entry.get())
            ts_forecast = TimeSeriesForecast(data)
            self.es_result_value = ts_forecast.exponential_smoothing(alpha, prior_forecast, observed_demand)

        # Exponential Smoothing
        es_df = pd.DataFrame({
            'Parameter': ['Smoothing Factor (alpha)', 'Prior Forecast', 'Observed Demand', 'Next Period Forecast'],
            'Value': [alpha, prior_forecast, observed_demand, self.es_result_value]
        })

        with pd.ExcelWriter(os.path.join(os.path.expanduser("~"), "Downloads", "forecast_results.xlsx"), engine='xlsxwriter') as writer:
            # Export Simple Moving Average
            sma_df.to_excel(writer, sheet_name='Simple Moving Average', index=False)
            sma_forecast_df.to_excel(writer, sheet_name='Simple Moving Average', startrow=len(sma_df) + 2, index=False)
            worksheet = writer.sheets['Simple Moving Average']
            worksheet.set_column('A:A', 20)  # Date
            worksheet.set_column('B:B', 20)  # Demand (Dt)
            worksheet.set_column('C:C', 30)  # Simple Moving Average Forecast

            # Export Weighted Moving Average
            wma_df.to_excel(writer, sheet_name='Weighted Moving Average', index=False)
            weights_df.to_excel(writer, sheet_name='Weighted Moving Average', startrow=len(wma_df) + 2, index=False)
            wma_forecast_df.to_excel(writer, sheet_name='Weighted Moving Average', startrow=len(wma_df) + len(weights_df) + 4, index=False)
            worksheet = writer.sheets['Weighted Moving Average']
            worksheet.set_column('A:A', 20)  # Date
            worksheet.set_column('B:B', 20)  # Demand (Dt)
            worksheet.set_column('C:D', 15)  # Weights
            worksheet.set_column('E:E', 30)  # Weighted Moving Average Forecast

            # Export Exponential Smoothing
            es_df.to_excel(writer, sheet_name='Exponential Smoothing', index=False)
            worksheet = writer.sheets['Exponential Smoothing']
            worksheet.set_column('A:A', 30)  # Parameter
            worksheet.set_column('B:B', 30)  # Value

        messagebox.showinfo("Export Successful", "Results exported to forecast_results.xlsx")

    def export_to_excel(self, sheet_name, df1, df2, df3, filename):
        with pd.ExcelWriter(os.path.join(os.path.expanduser("~"), "Downloads", filename), engine='xlsxwriter') as writer:
            df1.to_excel(writer, sheet_name=sheet_name, index=False)
            if df2 is not None:
                df2.to_excel(writer, sheet_name=sheet_name, startrow=len(df1) + 2, index=False)
            if df3 is not None:
                df3.to_excel(writer, sheet_name=sheet_name, startrow=len(df1) + len(df2) + 4, index=False)
            
            # Apply fixed column widths
            worksheet = writer.sheets[sheet_name]
            if sheet_name == 'Simple Moving Average':
                worksheet.set_column('A:A', 20)  # Date
                worksheet.set_column('B:B', 20)  # Demand (Dt)
                worksheet.set_column('C:C', 30)  # Simple Moving Average Forecast

            elif sheet_name == 'Weighted Moving Average':
                worksheet.set_column('A:A', 20)  # Date
                worksheet.set_column('B:B', 20)  # Demand (Dt)
                worksheet.set_column('C:D', 15)  # Weights
                worksheet.set_column('E:E', 30)  # Weighted Moving Average Forecast

            elif sheet_name == 'Exponential Smoothing':
                worksheet.set_column('A:A', 30)  # Parameter
                worksheet.set_column('B:B', 30)  # Value
        
        messagebox.showinfo("Export Successful", f"Results exported to {filename}")
