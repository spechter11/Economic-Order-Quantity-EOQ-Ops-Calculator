import pandas as pd

class ForecastErrorProcessor:
    def __init__(self, data):
        """
        Initialize with a DataFrame containing Month-Year, Forecast (Ft), and Demand (Dt).
        """
        self.data = data

    def calculate_errors(self):
        self.data['Forecast Error (Et)'] = self.data['Demand (Dt)'] - self.data['Forecast (Ft)']
        self.data['|Et|'] = self.data['Forecast Error (Et)'].abs()
        self.data['|Et|/Dt'] = (self.data['|Et|'] / self.data['Demand (Dt)'])

    def calculate_statistics(self):
        average_forecast_error = self.data['Forecast Error (Et)'].mean()
        mad = self.data['|Et|'].mean()
        mape = self.data['|Et|/Dt'].mean()
        return average_forecast_error, mad, mape

    def generate_results_table(self):
        self.calculate_errors()
        average_forecast_error, mad, mape = self.calculate_statistics()
        results = {
            'Parameter': ['Average Forecast Error', 'MAD', 'MAPE'],
            'Value': [average_forecast_error, mad, mape]
        }
        results_df = pd.DataFrame(results)
        return results_df

    def export_to_excel(self, filename="forecast_error_results.xlsx"):
        self.calculate_errors()
        results_df = self.generate_results_table()
        
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            self.data.to_excel(writer, sheet_name='Forecast Errors', index=False)
            results_df.to_excel(writer, sheet_name='Statistics', index=False)

            workbook = writer.book
            forecast_errors_sheet = writer.sheets['Forecast Errors']
            statistics_sheet = writer.sheets['Statistics']
            
            # Set the format for percentages and general number format
            percentage_format = workbook.add_format({'num_format': '0.00%'})
            general_format = workbook.add_format({'num_format': '0.00'})

            # Apply formatting to the 'Forecast Errors' sheet
            forecast_errors_sheet.set_column('A:A', 20)  # Month-Year
            forecast_errors_sheet.set_column('B:B', 15)  # Forecast (Ft)
            forecast_errors_sheet.set_column('C:C', 15)  # Demand (Dt)
            forecast_errors_sheet.set_column('D:D', 20, general_format)  # Forecast Error (Et)
            forecast_errors_sheet.set_column('E:E', 10, general_format)  # |Et|
            forecast_errors_sheet.set_column('F:F', 10, percentage_format)  # |Et|/Dt
            
            # Apply formatting to the 'Statistics' sheet
            statistics_sheet.set_column('A:A', 25)  # Parameter
            statistics_sheet.set_column('B:B', 15, general_format)  # Value
            # Specifically format MAPE as percentage
            for row_num, param in enumerate(results_df['Parameter']):
                if param == 'MAPE':
                    statistics_sheet.write(row_num + 1, 1, results_df.iloc[row_num, 1], percentage_format)
                else:
                    statistics_sheet.write(row_num + 1, 1, results_df.iloc[row_num, 1], general_format)

        print(f"Results exported to {filename}")