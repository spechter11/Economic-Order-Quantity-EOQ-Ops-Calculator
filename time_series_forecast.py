import pandas as pd
import numpy as np

class TimeSeriesForecast:
    def __init__(self, data):
        self.data = pd.Series(data)
    
    def simple_moving_average(self, window):
        if len(self.data) < window:
            raise ValueError("The length of the data must be greater than the window size.")
        return self.data[-window:].mean()

    def weighted_moving_average(self, weights):
        if len(weights) != len(self.data[-len(weights):]):
            raise ValueError("The length of the weights must be equal to the length of the data window.")
        return np.dot(self.data[-len(weights):], weights)

    def exponential_smoothing(self, alpha, prior_forecast, observed_demand):
        return alpha * observed_demand + (1 - alpha) * prior_forecast