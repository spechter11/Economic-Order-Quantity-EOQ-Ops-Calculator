from numpy import linspace
from pandas import DataFrame, ExcelWriter
import matplotlib.pyplot as plt
from eop_calculations import EOQCalculator
import pandas as pd

class EOQProcessor:
    def __init__(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, ordering_cost=None, standard_deviation=None, lead_time=None, service_level=None, weeks_per_year=50, EOQ=None):
        self.calculator = EOQCalculator(demand_rate, demand_yearly, purchase_cost, holding_cost_rate, ordering_cost, standard_deviation, lead_time, service_level, weeks_per_year, EOQ)

    def generate_input_table(self):
        inputs = {
            "Parameter": ["Demand Rate (units/week)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Ordering Cost (dollars/order)", "Standard Deviation (units/week)", 
                          "Lead Time (weeks)", "Service Level (%)", "Weeks per Year", "EOQ"],
            "Value": [self.calculator.demand_rate, self.calculator.demand_yearly, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.ordering_cost, self.calculator.standard_deviation, 
                      self.calculator.lead_time, self.calculator.service_level * 100 if self.calculator.service_level else None, 
                      self.calculator.weeks_per_year, self.calculator.EOQ]
        }
        input_df = pd.DataFrame(inputs)
        return input_df
    
    def generate_results_table(self):
        if self.calculator.EOQ is None:
            self.calculator.calculate_eoq()

        # Print intermediate values for debugging
        holding_cost = self.calculator.annual_holding_cost()
        ordering_cost = self.calculator.annual_ordering_cost()
        safety_stock = self.calculator.calculate_safety_stock()
        safety_stock_cost = self.calculator.annual_safety_stock_holding_cost()
        total_cost = self.calculator.total_annual_cost()
        TBO = self.calculator.time_between_orders()
        ROP = self.calculator.calculate_rop()
        EOQ = self.calculator.EOQ

        print(f"Holding Cost: {holding_cost}")
        print(f"Ordering Cost: {ordering_cost}")
        print(f"Safety Stock: {safety_stock}")
        print(f"Safety Stock Cost: {safety_stock_cost}")
        print(f"Total Cost: {total_cost}")
        print(f"Time Between Orders: {TBO}")
        print(f"Reorder Point: {ROP}")
        print(f"EOQ: {EOQ}")

        results = {
            "Parameter": ["Demand Rate (units/week)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Ordering Cost (dollars/order)", "Weeks per Year", 
                          "EOQ (units)", "Annual Holding Cost (dollars)", "Annual Ordering Cost (dollars)", 
                          "Annual Safety Stock Cost (dollars)", "Total Annual Cost (dollars)", 
                          "Safety Stock (units)", "Time Between Orders (weeks)", "Reorder Point (ROP)"],
            "Value": [self.calculator.demand_rate, self.calculator.D, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.ordering_cost, self.calculator.weeks_per_year, 
                      EOQ, holding_cost, ordering_cost, safety_stock_cost, total_cost, safety_stock, TBO, ROP],
            "Calculation": ["Given" if self.calculator.demand_rate is not None else "D / Weeks per Year", 
                            "Given" if self.calculator.demand_yearly is not None else "Demand Rate * Weeks per Year", 
                            "Given" if self.calculator.purchase_cost is not None else "Annual Holding Cost / Holding Cost Rate", 
                            "Given" if self.calculator.holding_cost_rate is not None else "Annual Holding Cost / Purchase Cost", 
                            "Given" if self.calculator.ordering_cost is not None else "(EOQ^2 * Annual Holding Cost) / (2 * Demand Yearly)",
                            "Given", 
                            "sqrt((2 * Demand Yearly * Ordering Cost) / Annual Holding Cost)", 
                            "(EOQ / 2) * Annual Holding Cost", 
                            "(Demand Yearly / EOQ) * Ordering Cost", 
                            "Safety Stock * Annual Holding Cost", 
                            "Annual Holding Cost + Annual Ordering Cost + Annual Safety Stock Cost",
                            "z * sigma_dLT", 
                            "EOQ / Demand Rate", 
                            "Demand Rate * Lead Time + Safety Stock"]
        }
        results_df = pd.DataFrame(results)
        return results_df
    
    def plot_costs(self, save_path=None):
        """
        Plot the EOQ model costs.
        
        Parameters:
        save_path (str): Path to save the plot image. If None, display the plot.
        """
        if self.calculator.EOQ is None:
            self.calculator.calculate_eoq()
        
        # Plot costs
        Q_range = linspace(1, 2 * self.calculator.EOQ, 500)
        holding_costs = (Q_range / 2) * self.calculator.H
        ordering_costs = (self.calculator.D / Q_range) * self.calculator.ordering_cost
        total_costs = holding_costs + ordering_costs

        plt.figure(figsize=(10, 6))
        plt.plot(Q_range, holding_costs, label='Annual Holding Cost (Q/2 * H)', color='green')
        plt.plot(Q_range, ordering_costs, label='Annual Ordering Cost (D/Q * S)', color='blue')
        plt.plot(Q_range, total_costs, label='Total Annual Cost', color='red')
        plt.axvline(self.calculator.EOQ, color='purple', linestyle='--', label=f'EOQ = {self.calculator.EOQ}')
        plt.axhline(min(total_costs), color='orange', linestyle='--', label=f'Minimum Total Cost = ${min(total_costs):.2f}')
        plt.xlabel('Order Quantity (Q)')
        plt.ylabel('Cost ($)')
        plt.title('EOQ Model: Costs vs. Order Quantity')
        plt.legend()
        plt.grid(True)

        if save_path:
            plt.savefig(save_path)
        else:
            plt.show()

    def export_to_excel(self, filename="eoq_results.xlsx"):
        """
        Export the input parameters and calculated results to an Excel file.
        
        Parameters:
        filename (str): The filename for the Excel file.
        """
        input_df = self.generate_input_table()
        results_df = self.generate_results_table()
        
        with ExcelWriter(filename, engine='xlsxwriter') as writer:
            input_df.to_excel(writer, sheet_name='Inputs', index=False)
            results_df.to_excel(writer, sheet_name='Results', index=False)
            
            workbook  = writer.book
            worksheet = writer.sheets['Results']
            
            # Insert the plot into the results sheet
            plot_path = 'eoq_plot.png'
            self.plot_costs(save_path=plot_path)
            worksheet.insert_image('D2', plot_path)
        
        print(f"Results exported to {filename}")

# Example usage
def main():
    demand_rate = 18  # units per week
    demand_yearly = None  # units per year, not provided
    purchase_cost = 60  # dollars per unit
    holding_cost_rate = 0.25  # annual holding cost rate
    ordering_cost = 45  # dollars per order
    standard_deviation = 5  # units per week
    lead_time = 2  # weeks
    service_level = 0.9  # 90% service level
    weeks_per_year = 52  # number of weeks in a year
    EOQ = None  # EOQ is to be calculated

    eoq_processor = EOQProcessor(
        demand_rate=demand_rate,
        demand_yearly=demand_yearly,
        purchase_cost=purchase_cost,
        holding_cost_rate=holding_cost_rate,
        ordering_cost=ordering_cost,
        standard_deviation=standard_deviation,
        lead_time=lead_time,
        service_level=service_level,
        weeks_per_year=weeks_per_year,
        EOQ=EOQ
    )

    eoq_processor.export_to_excel()

if __name__ == "__main__":
    main()