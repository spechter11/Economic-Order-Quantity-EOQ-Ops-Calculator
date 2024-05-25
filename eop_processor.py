from numpy import linspace
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

    def generate_input_table(self):
        inputs = {
            "Parameter": ["Demand Rate (units/day)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Holding Cost per Unit (dollars/unit/year)", "Ordering Cost (dollars/order)", 
                          "Standard Deviation (units/day)", "Lead Time (days)", "Service Level (%)", "Weeks per Year", "Days per Year", "EOQ", "Toggle Holding Stock"],
            "Value": [self.calculator.demand_rate, self.calculator.demand_yearly, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.holding_cost_per_unit, self.calculator.ordering_cost, 
                      self.calculator.standard_deviation_per_day, self.calculator.lead_time_days, self.calculator.service_level * 100 if self.calculator.service_level else None, 
                      self.calculator.weeks_per_year, self.calculator.days_per_year, self.calculator.EOQ, self.calculator.toggle_holding_stock]
        }
        input_df = pd.DataFrame(inputs)
        return input_df
    
    def generate_results_table(self):
        if self.calculator.EOQ is None:
            self.calculator.calculate_eoq()

        holding_cost = self.calculator.annual_holding_cost()
        ordering_cost = self.calculator.annual_ordering_cost()
        safety_stock = self.calculator.calculate_safety_stock() if self.calculator.toggle_holding_stock else None
        safety_stock_cost = self.calculator.annual_safety_stock_holding_cost() if self.calculator.toggle_holding_stock else None
        total_cost = self.calculator.total_annual_cost()
        TBO = self.calculator.time_between_orders()
        ROP = self.calculator.calculate_rop() if self.calculator.toggle_holding_stock else None
        EOQ = self.calculator.EOQ
        orders_per_year = self.calculator.number_of_orders_per_year()

        results = {
            "Parameter": ["Demand Rate (units/day)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Holding Cost per Unit (dollars/unit/year)", "Ordering Cost (dollars/order)", 
                          "Weeks per Year", "Days per Year", "EOQ (units)", "Annual Holding Cost (dollars)", 
                          "Annual Ordering Cost (dollars)", "Annual Safety Stock Cost (dollars)", 
                          "Total Annual Cost (dollars)", "Safety Stock (units)", 
                          "Time Between Orders (days)", "Reorder Point (ROP)", "Number of Orders per Year"],
            "Value": [self.calculator.demand_rate, self.calculator.D, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.holding_cost_per_unit, self.calculator.ordering_cost, 
                      self.calculator.weeks_per_year, self.calculator.days_per_year, EOQ, holding_cost, 
                      ordering_cost, safety_stock_cost, total_cost, safety_stock, TBO, ROP, orders_per_year],
            "Calculation": ["Given" if self.calculator.demand_rate is not None else "D / Days per Year", 
                            "Given" if self.calculator.demand_yearly is not None else "Demand Rate * Days per Year", 
                            "Given" if self.calculator.purchase_cost is not None else "Annual Holding Cost / Holding Cost Rate", 
                            "Given" if self.calculator.holding_cost_rate is not None else "Annual Holding Cost / Purchase Cost", 
                            "Given" if self.calculator.holding_cost_per_unit is not None else "Holding Cost Rate * Purchase Cost",
                            "Given" if self.calculator.ordering_cost is not None else "(EOQ^2 * Annual Holding Cost) / (2 * Demand Yearly)",
                            "Given", 
                            "Given",
                            "sqrt((2 * Demand Yearly * Ordering Cost) / Annual Holding Cost)", 
                            "(EOQ / 2) * Annual Holding Cost", 
                            "(Demand Yearly / EOQ) * Ordering Cost", 
                            "Safety Stock * Annual Holding Cost" if self.calculator.toggle_holding_stock else "Not Applicable",
                            "Annual Holding Cost + Annual Ordering Cost + Annual Safety Stock Cost",
                            "z * sigma_dLT" if self.calculator.toggle_holding_stock else "Not Applicable", 
                            "EOQ / Demand Rate", 
                            "Demand Rate * Lead Time + Safety Stock" if self.calculator.toggle_holding_stock else "Not Applicable",
                            "Demand Yearly / EOQ"]
        }
        results_df = pd.DataFrame(results)
        return results_df
    
    def plot_costs(self, save_path=None):
        if self.calculator.EOQ is None:
            self.calculator.calculate_eoq()
        
        Q_range = linspace(1, 2 * self.calculator.EOQ, 500)
        holding_costs = (Q_range / 2) * self.calculator.H if self.calculator.toggle_holding_stock else 0
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
            
            workbook  = writer.book
            worksheet = writer.sheets['Results']
            
            plot_path = 'eoq_plot.png'
            self.plot_costs(save_path=plot_path)
            worksheet.insert_image('D2', plot_path)
        
        print(f"Results exported to {filename}")

# Example usage
def main():
    demand_rate_per_day = 2.57  # units per day
    purchase_cost = 60  # dollars per unit
    holding_cost_per_unit = 15  # dollars per unit per year
    ordering_cost = 45  # dollars per order
    standard_deviation_per_day = 0.71  # units per day
    lead_time_days = 14  # days
    service_level = 0.9  # 90% service level
    days_per_year = 365  # number of days in a year
    toggle_holding_stock = False  # toggle holding stock

    eoq_processor = EOQProcessor(
        demand_rate=demand_rate_per_day,
        purchase_cost=purchase_cost,
        holding_cost_per_unit=holding_cost_per_unit,
        ordering_cost=ordering_cost,
        standard_deviation=standard_deviation_per_day,
        lead_time_days=lead_time_days,
        service_level=service_level,
        days_per_year=days_per_year,
        toggle_holding_stock=toggle_holding_stock
    )

    eoq_processor.export_to_excel()

if __name__ == "__main__":
    main()
