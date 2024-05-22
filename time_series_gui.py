import tkinter as tk
from tkinter.ttk import Label, Entry, Button

class ForecastApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Time Series Forecast")

        # Input fields for data
        ttk.Label(root, text="Enter time series data (comma separated values):").grid(column=0, row=0, sticky=tk.W, padx=10, pady=5)
        self.data_entry = tk.Entry(root, width=50)
        self.data_entry.grid(column=1, row=0, padx=10, pady=5)

        # Simple Moving Average
        ttk.Label(root, text="Simple Moving Average: Enter the window size (number of periods):").grid(column=0, row=1, sticky=tk.W, padx=10, pady=5)
        self.sma_window_entry = tk.Entry(root, width=10)
        self.sma_window_entry.grid(column=1, row=1, padx=10, pady=5)
        self.sma_result = ttk.Label(root, text="Result will appear here")
        self.sma_result.grid(column=2, row=1, padx=10, pady=5)
        ttk.Button(root, text="Calculate Simple Moving Average", command=self.calculate_sma).grid(column=3, row=1, padx=10, pady=5)

        # Weighted Moving Average
        ttk.Label(root, text="Weighted Moving Average: Enter weights (comma separated, sum should be 1):").grid(column=0, row=2, sticky=tk.W, padx=10, pady=5)
        self.wma_weights_entry = tk.Entry(root, width=50)
        self.wma_weights_entry.grid(column=1, row=2, padx=10, pady=5)
        self.wma_result = ttk.Label(root, text="Result will appear here")
        self.wma_result.grid(column=2, row=2, padx=10, pady=5)
        ttk.Button(root, text="Calculate Weighted Moving Average", command=self.calculate_wma).grid(column=3, row=2, padx=10, pady=5)

        # Exponential Smoothing
        ttk.Label(root, text="Exponential Smoothing: Enter the smoothing factor (alpha):").grid(column=0, row=3, sticky=tk.W, padx=10, pady=5)
        self.es_alpha_entry = tk.Entry(root, width=10)
        self.es_alpha_entry.grid(column=1, row=3, padx=10, pady=5)
        ttk.Label(root, text="Enter the prior forecast value:").grid(column=0, row=4, sticky=tk.W, padx=10, pady=5)
        self.es_prior_entry = tk.Entry(root, width=10)
        self.es_prior_entry.grid(column=1, row=4, padx=10, pady=5)
        ttk.Label(root, text="Enter the observed demand for the last period:").grid(column=0, row=5, sticky=tk.W, padx=10, pady=5)
        self.es_observed_entry = tk.Entry(root, width=10)
        self.es_observed_entry.grid(column=1, row=5, padx=10, pady=5)
        self.es_result = ttk.Label(root, text="Result will appear here")
        self.es_result.grid(column=2, row=5, padx=10, pady=5)
        ttk.Button(root, text="Calculate Exponential Smoothing", command=self.calculate_es).grid(column=3, row=5, padx=10, pady=5)

    def get_data(self):
        data_str = self.data_entry.get()
        return [float(i) for i in data_str.split(",")]

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

root = tk.Tk()
app = ForecastApp(root)
root.mainloop()
