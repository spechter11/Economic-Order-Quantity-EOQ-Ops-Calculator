from numpy import linspace
from pandas import DataFrame, ExcelWriter
from matplotlib.pyplot import plot, show, figure, axvline, axhline, xlabel, ylabel, title, legend, grid
from eop_calculations import EOQCalculator


class EOQProcessor:
    def __init__(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, ordering_cost=None, weeks_per_year=50, EOQ=None):
        self.calculator = EOQCalculator(demand_rate, demand_yearly, purchase_cost, holding_cost_rate, ordering_cost, weeks_per_year, EOQ)
    
    def generate_input_table(self):
        inputs = {
            "Parameter": ["Demand Rate (units/week)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Ordering Cost (dollars/order)", "Weeks per Year", "EOQ"],
            "Value": [self.calculator.demand_rate, self.calculator.demand_yearly, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.ordering_cost, self.calculator.weeks_per_year, self.calculator.EOQ]
        }
        input_df = pd.DataFrame(inputs)
        return input_df
    
    def generate_results_table(self):
        try:
            if self.calculator.EOQ is None:
                self.calculator.calculate_eoq()
            holding_cost = round(self.calculator.annual_holding_cost(), 2)
            ordering_cost = round(self.calculator.annual_ordering_cost(), 2)
            TBO = round(self.calculator.time_between_orders(), 2)
            EOQ = round(self.calculator.EOQ, 2)
        except ValueError as e:
            holding_cost = None
            ordering_cost = None
            TBO = None
            EOQ = None

        results = {
            "Parameter": ["Demand Rate (units/week)", "Demand Yearly (units/year)", "Purchase Cost (dollars/unit)", 
                          "Holding Cost Rate (annual %)", "Ordering Cost (dollars/order)", "Weeks per Year", 
                          "EOQ (units)", "Annual Holding Cost (dollars)", "Annual Ordering Cost (dollars)", 
                          "Time Between Orders (weeks)"],
            "Value": [self.calculator.demand_rate, self.calculator.D, self.calculator.purchase_cost, 
                      self.calculator.holding_cost_rate, self.calculator.ordering_cost, self.calculator.weeks_per_year, 
                      EOQ, holding_cost, ordering_cost, TBO],
            "Calculation": ["Given" if self.calculator.demand_rate is not None else "D / Weeks per Year", 
                            "Given" if self.calculator.demand_yearly is not None else "Demand Rate * Weeks per Year", 
                            "Given" if self.calculator.purchase_cost is not None else "Annual Holding Cost / Holding Cost Rate", 
                            "Given" if self.calculator.holding_cost_rate is not None else "Annual Holding Cost / Purchase Cost", 
                            "Given" if self.calculator.ordering_cost is not None else "(EOQ^2 * Annual Holding Cost) / (2 * Demand Yearly)",
                            "Given", 
                            "sqrt((2 * Demand Yearly * Ordering Cost) / Annual Holding Cost)", 
                            "(EOQ / 2) * Annual Holding Cost", 
                            "(Demand Yearly / EOQ) * Ordering Cost", 
                            "EOQ / Demand Rate"]
        }
        results_df = pd.DataFrame(results)
        return results_df
    
    def plot_costs(self, save_path=None):
        if self.calculator.EOQ is None:
            self.calculator.calculate_eoq()
        
        # Plot costs
        Q_range = np.linspace(1, 2 * self.calculator.EOQ, 500)
        holding_costs = (Q_range / 2) * self.calculator.H
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
        
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            input_df.to_excel(writer, sheet_name='Inputs', index=False)
            results_df.to_excel(writer, sheet_name='Results', index=False)
            
            workbook  = writer.book
            worksheet = writer.sheets['Results']
            
            # Insert the plot into the results sheet
            plot_path = 'eoq_plot.png'
            self.plot_costs(save_path=plot_path)
            worksheet.insert_image('D2', plot_path)
        
        print(f"Results exported to {filename}")
        print(f"Results exported to {filename}")
