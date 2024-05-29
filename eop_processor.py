from numpy import linspace
import numpy as np
from pandas import DataFrame, ExcelWriter
import matplotlib.pyplot as plt
from eop_calculations import EOQCalculator
import pandas as pd

class EOQProcessor:
    def __init__(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, holding_cost_per_unit=None, ordering_cost=None, standard_deviation=None, standard_deviation_per_day=None, lead_time=None, lead_time_days=None, service_level=None, weeks_per_year=52, days_per_year=365, EOQ=None, toggle_holding_stock=True):
        self.calculator = EOQCalculator(
            demand_rate=demand_rate,
            demand_yearly=demand_yearly,
            purchase_cost=purchase_cost,
            holding_cost_rate=holding_cost_rate,
            holding_cost_per_unit=holding_cost_per_unit,
            ordering_cost=ordering_cost,
            standard_deviation=standard_deviation,
            standard_deviation_per_day=standard_deviation_per_day,
            lead_time=lead_time,
            lead_time_days=lead_time_days,
            service_level=service_level,
            weeks_per_year=weeks_per_year,
            days_per_year=days_per_year,
            EOQ=EOQ,
            toggle_holding_stock=toggle_holding_stock
        )

        # Ensure EOQ is calculated
        self.calculate_eoq()

    def calculate_eoq(self):
        try:
            self.calculator.EOQ = self.calculator.calculate_eoq()
        except ValueError as e:
            raise ValueError("EOQ calculation failed: " + str(e))

    def generate_input_table(self):
        inputs = {
            "Parameter": ["Demand Rate (units/day)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)",
                          "Holding Cost Rate (annual %)", "Holding Cost per Unit (dollars/unit/year)", "Ordering Cost (dollars/order)", "Standard Deviation (units/week)",
                          "Standard Deviation (units/day)", "Lead Time (days)", "Service Level (%)", "Weeks per Year", "Days per Year", "EOQ", "Toggle Holding Stock", "Z Score"],
            "Value": [self.calculator.demand_rate, self.calculator.demand_yearly, self.calculator.purchase_cost,
                      self.calculator.holding_cost_rate, self.calculator.holding_cost_per_unit, self.calculator.ordering_cost,
                      self.calculator.standard_deviation_per_day, self.calculator.standard_deviation, self.calculator.lead_time_days, self.calculator.service_level * 100 if self.calculator.service_level else None,
                      self.calculator.weeks_per_year, self.calculator.days_per_year, self.calculator.EOQ, self.calculator.toggle_holding_stock, self.calculator.z]
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
        orders_per_year = self.calculator.number_of_orders_per_year()

        # Print intermediate values for debugging
        print(f"Holding Cost: {holding_cost}")
        print(f"Ordering Cost: {ordering_cost}")
        print(f"Safety Stock: {safety_stock}")
        print(f"Safety Stock Cost: {safety_stock_cost}")
        print(f"Total Cost: {total_cost}")
        print(f"Time Between Orders: {TBO}")
        print(f"Reorder Point: {ROP}")
        print(f"EOQ: {EOQ}")

        results = {
            "Parameter": ["Demand Rate (units/day)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Holding Cost per Unit (dollars/unit/year)", "Ordering Cost (dollars/order)", 
                          "Weeks per Year", "Days per Year", "EOQ (units)", "Annual Holding Cost (dollars)", 
                          "Annual Ordering Cost (dollars)", "Annual Safety Stock Cost (dollars)", 
                          "Total Annual Cost (dollars)", "Safety Stock (units)", 
                          "Time Between Orders (days)", "Reorder Point (ROP)", "Number of Orders per Year"],
            "Value": [self.calculator.demand_rate, self.calculator.demand_yearly, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.holding_cost_per_unit, self.calculator.ordering_cost, 
                      self.calculator.weeks_per_year, self.calculator.days_per_year, EOQ, holding_cost, 
                      ordering_cost, safety_stock_cost, total_cost, safety_stock, TBO, ROP, orders_per_year],
            "Calculation": ["Given" if self.calculator.demand_rate is not None else "Demand Yearly / Days per Year", 
                            "Given" if self.calculator.demand_yearly is not None else "Demand Rate * Days per Year", 
                            "Given" if self.calculator.purchase_cost is not None else "N/A", 
                            "Given" if self.calculator.holding_cost_rate is not None else "Annual Holding Cost / Purchase Cost", 
                            "Given" if self.calculator.holding_cost_per_unit is not None else "Holding Cost Rate * Purchase Cost",
                            "Given" if self.calculator.ordering_cost is not None else "(EOQ^2 * Annual Holding Cost) / (2 * Demand Yearly)",
                            "Given", 
                            "Given",
                            "sqrt((2 * Demand Yearly * Ordering Cost) / (Holding Cost per Unit))", 
                            "(EOQ / 2) * Holding Cost per Unit", 
                            "(Demand Yearly / EOQ) * Ordering Cost", 
                            "Safety Stock * Holding Cost per Unit" if self.calculator.toggle_holding_stock else "Not Applicable",
                            "Annual Holding Cost + Annual Ordering Cost + (Annual Safety Stock Cost if applicable)",
                            "z * sigma_dLT" if self.calculator.toggle_holding_stock else "Not Applicable", 
                            "EOQ / Demand Rate", 
                            "Demand Rate * Lead Time + Safety Stock" if self.calculator.toggle_holding_stock else "Not Applicable",
                            "Demand Yearly / EOQ"]
        }
        
        results_df = pd.DataFrame(results)
        return results_df

    def plot_costs(self, save_path=None):
        # Ensure EOQ is calculated and non-zero before plotting costs
        if self.calculator.EOQ is None or self.calculator.EOQ == 0:
            raise ValueError("EOQ must be calculated and non-zero before plotting costs")

        Q_range = linspace(1, 2 * self.calculator.EOQ, 500)
        if self.calculator.toggle_holding_stock:
            holding_costs = (Q_range / 2) * self.calculator.H
        else:
            holding_costs = np.zeros_like(Q_range)
        
        ordering_costs = (self.calculator.D / Q_range) * self.calculator.ordering_cost
        total_costs = holding_costs + ordering_costs

        plt.figure(figsize=(10, 6))
        plt.plot(Q_range, holding_costs, label='Annual Holding Cost (Q/2 * H)', color='green')
        plt.plot(Q_range, ordering_costs, label='Annual Ordering Cost (D/Q * S)', color='blue')
        plt.plot(Q_range, total_costs, label='Total Annual Cost', color='red')
        plt.axvline(self.calculator.EOQ, color='purple', linestyle='--', label=f'EOQ = {self.calculator.EOQ:.2f}')
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
        input_df = self.generate_input_table()
        results_df = self.generate_results_table()
        
        with ExcelWriter(filename, engine='xlsxwriter') as writer:
            input_df.to_excel(writer, sheet_name='Inputs', index=False)
            results_df.to_excel(writer, sheet_name='Results', index=False)
            
            workbook = writer.book
            input_sheet = writer.sheets['Inputs']
            results_sheet = writer.sheets['Results']
            
            # Formatting
            general_format = workbook.add_format({'num_format': '0.00'})
            percentage_format = workbook.add_format({'num_format': '0%'})
            
            # Format columns in the Inputs sheet
            input_sheet.set_column('A:A', 30)  # Parameter column
            input_sheet.set_column('B:B', 20)  # Value column
            
            # Format columns in the Results sheet
            results_sheet.set_column('A:A', 40)  # Parameter column
            results_sheet.set_column('B:B', 20, general_format)  # Value column
            results_sheet.set_column('C:C', 40)  # Calculation column

            # Insert the plot into the results sheet
            plot_path = 'eoq_plot.png'
            self.plot_costs(save_path=plot_path)
            results_sheet.insert_image('D2', plot_path)
        
        print(f"Results exported to {filename}")

def main():
    demand_rate_per_day = 15.0  # units per day
    demand_yearly = None
    purchase_cost = 11.7  # dollars per unit
    holding_cost_rate = .28
    holding_cost_per_unit = None  # dollars per unit per year
    ordering_cost = 54.0  # dollars per order
    standard_deviation_per_day = 6.124  # units per day
    lead_time_days = 18.0  # days
    service_level = 0.8  # 80% service level
    weeks_per_year = None  # Default value
    days_per_year = 312  # number of days in a year
    EOQ = None
    toggle_holding_stock = True  # toggle holding stock

    eoq_processor = EOQProcessor(
        demand_rate=demand_rate_per_day,
        demand_yearly=demand_yearly,
        purchase_cost=purchase_cost,
        holding_cost_rate=holding_cost_rate,
        holding_cost_per_unit=holding_cost_per_unit,
        ordering_cost=ordering_cost,
        standard_deviation=standard_deviation_per_day,
        lead_time_days=lead_time_days,
        service_level=service_level,
        weeks_per_year=weeks_per_year,
        days_per_year=days_per_year,
        EOQ=EOQ,
        toggle_holding_stock=toggle_holding_stock
    )

    try:
        eoq_processor.export_to_excel()
        print("Results successfully exported to Excel.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
