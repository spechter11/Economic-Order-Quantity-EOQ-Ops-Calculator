
# EOQ and Time Series Forecasting Application

This project provides a comprehensive application for calculating the Economic Order Quantity (EOQ), performing time series forecasting, and calculating forecast errors. The application is built using Python and Tkinter for the GUI, and it includes functionality for Simple Moving Average (SMA), Weighted Moving Average (WMA), and Exponential Smoothing (ES) forecasts.

## Features

- **EOQ Calculator**: Calculates Economic Order Quantity, annual holding cost, annual ordering cost, and time between orders.
- **Time Series Forecasting**: Provides forecasting using Simple Moving Average (SMA), Weighted Moving Average (WMA), and Exponential Smoothing (ES) methods.
- **Forecast Error Calculation**: Calculates forecast errors including Mean Absolute Deviation (MAD) and Mean Absolute Percentage Error (MAPE).
- **Visualization**: Plots EOQ model costs and provides visual representation of forecast results.
- **Export to Excel**: Exports results and plots to Excel files for further analysis.

## Installation

To run this project, ensure you have Python installed along with the required packages. You can install the necessary packages using `pip`:

```bash
pip install numpy pandas matplotlib tkinter
```

## Usage

### Running the Application

To start the application, run the `main.py` file:

```bash
python main.py
```

### EOQ Calculator

1. **Inputs**:
   - Demand Rate (units/week)
   - Demand Yearly (units/year)
   - Purchase Cost (dollars/unit)
   - Holding Cost Rate (annual %)
   - Ordering Cost (dollars/order)
   - Weeks per Year
   - EOQ (optional)

2. **Actions**:
   - **Calculate**: Calculates EOQ and associated parameters.
   - **Visualize**: Plots EOQ model costs.
   - **Export**: Exports results to an Excel file.

### Time Series Forecasting

1. **Inputs**:
   - Time series data (space-separated rows)
   - SMA window size (number of periods)
   - WMA weights (space-separated, sum should be 1)
   - ES parameters (smoothing factor, prior forecast value, observed demand for last period)

2. **Actions**:
   - **Calculate SMA/WMA/ES**: Calculates the respective forecasts.
   - **Export SMA/WMA/ES**: Exports the forecast results to an Excel file.

### Forecast Error Calculation

1. **Inputs**:
   - Data for error calculation (Month-Year, Forecast, Demand) - space separated rows

2. **Actions**:
   - **Calculate Errors**: Calculates forecast errors including Mean Absolute Deviation (MAD) and Mean Absolute Percentage Error (MAPE).
   - **Export**: Exports the error calculation results to an Excel file.
