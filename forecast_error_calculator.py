import pandas as pd

class ForecastErrorCalculator:
    def __init__(self, data):
        """
        Initialize with a DataFrame containing Month-Year, Forecast (Ft), and Demand (Dt).
        """
        self.data = data

    def calculate_errors(self):
        self.data['Forecast Error (Et)'] = self.data['Forecast (Ft)'] - self.data['Demand (Dt)']
        self.data['|Et|'] = self.data['Forecast Error (Et)'].abs()
        self.data['|Et|/Dt'] = (self.data['|Et|'] / self.data['Demand (Dt)']) * 100

    def calculate_statistics(self):
        average_forecast_error = self.data['Forecast Error (Et)'].mean()
        mad = self.data['|Et|'].mean()
        mape = self.data['|Et|/Dt'].mean()
        return average_forecast_error, mad, mape

    def get_results(self):
        self.calculate_errors()
        average_forecast_error, mad, mape = self.calculate_statistics()
        results = {
            'Average Forecast Error': average_forecast_error,
            'MAD': mad,
            'MAPE': mape
        }
        return self.data, results
