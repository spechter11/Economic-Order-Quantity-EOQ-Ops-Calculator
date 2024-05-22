import math
from numpy import linspace, sqrt
from matplotlib.pyplot import plot, show, figure, axvline, axhline, xlabel, ylabel, title, legend, grid

class EOQCalculator:
    def __init__(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, ordering_cost=None, weeks_per_year=50, EOQ=None):
        self.demand_rate = demand_rate  # units per week
        self.demand_yearly = demand_yearly  # units per year
        self.purchase_cost = purchase_cost  # cost per unit
        self.holding_cost_rate = holding_cost_rate  # holding cost rate per year
        self.ordering_cost = ordering_cost  # cost per order
        self.weeks_per_year = weeks_per_year  # number of weeks in a year
        self.EOQ = EOQ
        self.H = None
        self.D = None
        
        self.update_calculations()
    
    def update_calculations(self):
        if self.demand_rate is not None:
            self.D = self.demand_rate * self.weeks_per_year  # annual demand in units/year
        elif self.demand_yearly is not None:
            self.D = self.demand_yearly
            self.demand_rate = self.D / self.weeks_per_year  # calculate weekly demand from annual demand
        if self.holding_cost_rate is not None and self.purchase_cost is not None:
            self.H = self.holding_cost_rate * self.purchase_cost  # annual holding cost per unit
        if self.D is not None and self.ordering_cost is not None and self.H is not None and self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        if self.EOQ is not None:
            self.solve_missing_parameters()
    
    def solve_missing_parameters(self):
        if self.D is None and self.ordering_cost is not None and self.H is not None:
            self.D = (self.EOQ ** 2 * self.H) / (2 * self.ordering_cost)
            self.demand_rate = self.D / self.weeks_per_year
        if self.ordering_cost is None and self.D is not None and self.H is not None:
            self.ordering_cost = (self.EOQ ** 2 * self.H) / (2 * self.D)
        if self.H is None and self.D is not None and self.ordering_cost is not None:
            self.H = (2 * self.D * self.ordering_cost) / (self.EOQ ** 2)
        if self.holding_cost_rate is None and self.H is not None and self.purchase_cost is not None:
            self.holding_cost_rate = self.H / self.purchase_cost
        if self.purchase_cost is None and self.H is not None and self.holding_cost_rate is not None:
            self.purchase_cost = self.H / self.holding_cost_rate
        if self.demand_rate is None and self.D is not None and self.weeks_per_year is not None:
            self.demand_rate = self.D / self.weeks_per_year
    
    def set_parameters(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, ordering_cost=None, weeks_per_year=None, EOQ=None):
        """
        Set or update the parameters for EOQ calculation.
        """
        if demand_rate is not None:
            self.demand_rate = demand_rate
        if demand_yearly is not None:
            self.demand_yearly = demand_yearly
        if purchase_cost is not None:
            self.purchase_cost = purchase_cost
        if holding_cost_rate is not None:
            self.holding_cost_rate = holding_cost_rate
        if ordering_cost is not None:
            self.ordering_cost = ordering_cost
        if weeks_per_year is not None:
            self.weeks_per_year = weeks_per_year
        if EOQ is not None:
            self.EOQ = EOQ
        
        self.update_calculations()
    
    def calculate_eoq(self):
        """
        Calculate the Economic Order Quantity (EOQ).
        
        Returns:
        float: The calculated EOQ.
        """
        if self.D is None or self.ordering_cost is None or self.H is None:
            raise ValueError("Insufficient parameters to calculate EOQ")
        return math.sqrt((2 * self.D * self.ordering_cost) / self.H)
    
    def annual_holding_cost(self):
        """
        Calculate the annual total cycle inventory holding costs.
        
        Returns:
        float: The annual total cycle inventory holding costs.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        return (self.EOQ / 2) * self.H
    
    def annual_ordering_cost(self):
        """
        Calculate the annual total ordering costs.
        
        Returns:
        float: The annual total ordering costs.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        return (self.D / self.EOQ) * self.ordering_cost
    
    def time_between_orders(self):
        """
        Calculate the time between two orders (TBO).
        
        Returns:
        float: The time between two orders (TBO) in weeks.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        return self.EOQ / self.demand_rate
    
    def plot_costs(self):
        """
        Plot the EOQ model costs.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        
        # Range of order quantities to plot
        Q_range = np.linspace(1, 2 * self.EOQ, 500)

        # Calculate costs
        holding_costs = (Q_range / 2) * self.H
        ordering_costs = (self.D / Q_range) * self.ordering_cost
        total_costs = holding_costs + ordering_costs

        # Plot the costs
        plt.figure(figsize=(10, 6))
        plt.plot(Q_range, holding_costs, label='Annual Holding Cost (Q/2 * H)', color='green')
        plt.plot(Q_range, ordering_costs, label='Annual Ordering Cost (D/Q * S)', color='blue')
        plt.plot(Q_range, total_costs, label='Total Annual Cost', color='red')

        # Mark the EOQ
        plt.axvline(self.EOQ, color='purple', linestyle='--', label=f'EOQ = {self.EOQ:.2f}')
        plt.axhline(min(total_costs), color='orange', linestyle='--', label=f'Minimum Total Cost = ${min(total_costs):.2f}')

        # Add labels and title
        plt.xlabel('Order Quantity (Q)')
        plt.ylabel('Cost ($)')
        plt.title('EOQ Model: Costs vs. Order Quantity')
        plt.legend()
        plt.grid(True)
        plt.show()
    
    def print_results(self):
        """
        Print the EOQ and associated parameters and costs.
        """
        # Ensure that EOQ and necessary calculations are done
        try:
            if self.EOQ is None:
                self.calculate_eoq()
            holding_cost = self.annual_holding_cost()
            ordering_cost = self.annual_ordering_cost()
            TBO = self.time_between_orders()
        except ValueError as e:
            print(f"Error: {e}")
            holding_cost = None
            ordering_cost = None
            TBO = None
        
        print("EOQ Calculation Results:")
        print(f"Demand rate (units/week): {self.demand_rate if self.demand_rate is not None else 'None'}")
        print(f"Demand (units/year): {self.D if self.D is not None else 'None'}")
        print(f"Purchase cost (dollars/unit): {self.purchase_cost if self.purchase_cost is not None else 'None'}")
        print(f"Holding cost rate (annual %): {self.holding_cost_rate if self.holding_cost_rate is not None else 'None'}")
        print(f"Ordering cost (dollars/order): {self.ordering_cost if self.ordering_cost is not None else 'None'}")
        print(f"Weeks per year: {self.weeks_per_year if self.weeks_per_year is not None else 'None'}")
        print(f"Economic Order Quantity (EOQ): {self.EOQ:.2f} units" if self.EOQ is not None else "EOQ: None")
        print(f"Annual holding cost per unit (dollars/unit/year): {self.H if self.H is not None else 'None'}")
        print(f"The annual total cycle inventory holding costs are: ${holding_cost:.2f}" if holding_cost is not None else "Annual total cycle inventory holding costs: None")
        print(f"The annual total ordering costs are: ${ordering_cost:.2f}" if ordering_cost is not None else "Annual total ordering costs: None")
        print(f"The time between two orders (TBO) is: {TBO:.2f} weeks" if TBO is not None else "Time between two orders (TBO): None")