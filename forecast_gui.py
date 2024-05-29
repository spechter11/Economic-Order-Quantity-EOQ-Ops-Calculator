import tkinter as tk
from tkinter import ttk, messagebox
from pandas import DataFrame, ExcelWriter
import numpy as np
import os
from time_series_forecast import TimeSeriesForecast


class ForecastApp:
    def __init__(self, parent, scale):
        self.parent = parent
        self.scale = scale  # Adjust this scale factor as needed
        self.sma_result_value = None
        self.wma_result_value = None
        self.es_result_value = None
        self.create_scrollable_canvas(parent)
        self.create_forecast_widgets(self.scrollable_frame)
        self.configure_grid_weights(self.scrollable_frame)

    def create_scrollable_canvas(self, parent):
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.scrollable_frame = scrollable_frame

    def create_forecast_widgets(self, parent):
        # Adding a title
        title_label = ttk.Label(parent, text="Time Series Forecasting", font=("Helvetica", int(25 * self.scale), "bold"))
        title_label.grid(row=0, column=0, columnspan=5, pady=10, sticky="ew")

        # Instructions
        instructions = (
            "Enter time series data as space-separated values, one per line, e.g.:"
            "\nDate1 Value1"
            "\nDate2 Value2"
            "\n...\n\n"
            "Enter weights for WMA in the format 'w0 value0', 'w1 value1', ... , one per line, e.g.:"
            "\nw0 0.3"
            "\nw1 0.25"
            "\nw2 0.1"
            "\nw3 0.1"
            "\nw4 0.08"
            "\nw5 0.07"
            "\nw6 0.05"
            "\nw7 0.05"
        )
        instructions_label = ttk.Label(parent, text=instructions, justify=tk.LEFT, font=("Helvetica", int(15 * self.scale), "bold"))
        instructions_label.grid(row=1, column=0, columnspan=5, pady=5, sticky="ew", padx=10)

        # Input fields for data
        text_font = ("Helvetica", int(15 * self.scale))  # Define a larger font size for text widgets
        ttk.Label(parent, text="Enter time series data:", font=("Helvetica", int(15 * self.scale))).grid(column=0, row=2, sticky=tk.W, padx=10, pady=5)
        self.data_entry = tk.Text(parent, height=10, width=50, font=text_font)
        self.data_entry.grid(column=1, row=2, columnspan=4, padx=10, pady=5, sticky="ew")
        self.data_entry.insert(tk.END, "Date1 10\nDate2 15\nDate3 20\nDate4 25\nDate5 30\nDate6 35\nDate7 40\nDate8 45")  # Sample data

        # Simple Moving Average
        ttk.Label(parent, text="Simple Moving Average: Enter window size:", font=("Helvetica", int(15 * self.scale))).grid(column=0, row=3, sticky=tk.W, padx=10, pady=5)
        self.sma_window_entry = tk.Entry(parent, width=10, font=text_font)
        self.sma_window_entry.grid(column=1, row=3, padx=10, pady=5, sticky="ew")
        self.sma_result = ttk.Label(parent, text="Result will appear here", font=("Helvetica", int(15 * self.scale)))
        self.sma_result.grid(column=2, row=3, padx=10, pady=5, sticky="ew")
        ttk.Button(parent, text="Calculate SMA", command=self.calculate_sma, padding=(10, 5)).grid(column=3, row=3, padx=10, pady=5, sticky="ew")
        ttk.Button(parent, text="Export SMA to Excel", command=self.export_sma_to_excel, padding=(10, 5)).grid(column=4, row=3, padx=10, pady=5, sticky="ew")

        # Weighted Moving Average
        ttk.Label(parent, text="Weighted Moving Average: Enter weights:", font=("Helvetica", int(15 * self.scale))).grid(column=0, row=4, sticky=tk.W, padx=10, pady=5)
        self.wma_weights_entry = tk.Text(parent, height=10, width=50, font=text_font)
        self.wma_weights_entry.grid(column=1, row=4, columnspan=4, padx=10, pady=5, sticky="ew")
        self.wma_weights_entry.insert(tk.END, "w0 0.3\nw1 0.25\nw2 0.1\nw3 0.1\nw4 0.08\nw5 0.07\nw6 0.05\nw7 0.05")  # Sample weights
        self.wma_result = ttk.Label(parent, text="Result will appear here", font=("Helvetica", int(15 * self.scale)))
        self.wma_result.grid(column=2, row=5, padx=10, pady=5, sticky="ew")
        ttk.Button(parent, text="Calculate WMA", command=self.calculate_wma, padding=(10, 5)).grid(column=3, row=5, padx=10, pady=5, sticky="ew")
        ttk.Button(parent, text="Export WMA to Excel", command=self.export_wma_to_excel, padding=(10, 5)).grid(column=4, row=5, padx=10, pady=5, sticky="ew")

        # Exponential Smoothing
        ttk.Label(parent, text="Exponential Smoothing: Enter smoothing factor (alpha):", font=("Helvetica", int(15 * self.scale))).grid(column=0, row=6, sticky=tk.W, padx=10, pady=5)
        self.es_alpha_entry = tk.Entry(parent, width=10, font=text_font)
        self.es_alpha_entry.grid(column=1, row=6, padx=10, pady=5, sticky="ew")
        ttk.Label(parent, text="Enter prior forecast value:", font=("Helvetica", int(15 * self.scale))).grid(column=0, row=7, sticky=tk.W, padx=10, pady=5)
        self.es_prior_entry = tk.Entry(parent, width=10, font=text_font)
        self.es_prior_entry.grid(column=1, row=7, padx=10, pady=5, sticky="ew")
        ttk.Label(parent, text="Enter observed demand for last period:", font=("Helvetica", int(15 * self.scale))).grid(column=0, row=8, sticky=tk.W, padx=10, pady=5)
        self.es_observed_entry = tk.Entry(parent, width=10, font=text_font)
        self.es_observed_entry.grid(column=1, row=8, padx=10, pady=5, sticky="ew")
        self.es_result = ttk.Label(parent, text="Result will appear here", font=("Helvetica", int(15 * self.scale)))
        self.es_result.grid(column=2, row=8, padx=10, pady=5, sticky="ew")
        ttk.Button(parent, text="Calculate ES", command=self.calculate_es, padding=(10, 5)).grid(column=3, row=8, padx=10, pady=5, sticky="ew")
        ttk.Button(parent, text="Export ES to Excel", command=self.export_es_to_excel, padding=(10, 5)).grid(column=4, row=8, padx=10, pady=5, sticky="ew")

        # Export All
        ttk.Button(parent, text="Export All to Excel", command=self.export_all_to_excel).grid(column=0, row=9, columnspan=5, padx=10, pady=20, sticky="ew")

    def configure_grid_weights(self, parent):
        for i in range(10):
            parent.grid_rowconfigure(i, weight=1)
        for i in range(5):
            parent.grid_columnconfigure(i, weight=1)

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
        df = DataFrame({'Date': dates, 'Demand (Dt)': data})
        forecast_df = DataFrame({
            'Date': ['Next Period'],
            'Simple Moving Average Forecast': [self.sma_result_value]
        })
        self.export_to_excel('Simple Moving Average', df, forecast_df, None, 'sma_result.xlsx')

    def export_wma_to_excel(self):
        data = self.get_data()
        weights = self.get_weights()
        dates = [row.split()[0] for row in self.data_entry.get("1.0", tk.END).strip().split('\n')]
        df = DataFrame({'Date': dates, 'Demand (Dt)': data})
        weights_labels = [f'w{i}' for i in range(len(weights))]
        weights_df = DataFrame({'Weight #': weights_labels, 'Weight': weights})
        forecast_df = DataFrame({
            'Date': ['Next Period'],
            'Weighted Moving Average Forecast': [self.wma_result_value]
        })
        self.export_to_excel('Weighted Moving Average', df, weights_df, forecast_df, 'wma_result.xlsx')

    def export_es_to_excel(self):
        alpha = self.parse_number(self.es_alpha_entry.get())
        prior_forecast = self.parse_number(self.es_prior_entry.get())
        observed_demand = self.parse_number(self.es_observed_entry.get())
        es_df = DataFrame({
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
        sma_df = DataFrame({'Date': dates, 'Demand (Dt)': data})
        sma_forecast_df = DataFrame({
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
        weights_labels = [f'w{i}' for i in range(len(weights))]
        wma_df = DataFrame({'Date': dates, 'Demand (Dt)': data})
        weights_df = DataFrame({'Weight #': weights_labels, 'Weight': weights})
        wma_forecast_df = DataFrame({
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
        es_df = DataFrame({
            'Parameter': ['Smoothing Factor (alpha)', 'Prior Forecast', 'Observed Demand', 'Next Period Forecast'],
            'Value': [alpha, prior_forecast, observed_demand, self.es_result_value]
        })

        with ExcelWriter(os.path.join(os.path.expanduser("~"), "Downloads", "forecast_results.xlsx"), engine='xlsxwriter') as writer:
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
        with ExcelWriter(os.path.join(os.path.expanduser("~"), "Downloads", filename), engine='xlsxwriter') as writer:
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


# Setup and run the application
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Time Series Forecasting")
    app = ForecastApp(root)
    root.mainloop()
