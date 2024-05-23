import math
from numpy import linspace
from matplotlib.pyplot import plot, show, figure, axvline, axhline, xlabel, ylabel, title, legend, grid
from scipy.stats import norm

class EOQCalculator:
    def __init__(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, ordering_cost=None, standard_deviation=None, lead_time=None, service_level=None, weeks_per_year=50, EOQ=None):
        self.demand_rate = demand_rate  # units per week
        self.demand_yearly = demand_yearly  # units per year
        self.purchase_cost = purchase_cost  # cost per unit
        self.holding_cost_rate = holding_cost_rate  # holding cost rate per year
        self.ordering_cost = ordering_cost  # cost per order
        self.standard_deviation = standard_deviation  # standard deviation of demand per week
        self.lead_time = lead_time  # lead time in weeks
        self.service_level = service_level  # service level as a decimal (e.g., 0.9 for 90%)
        self.weeks_per_year = weeks_per_year  # number of weeks in a year
        self.EOQ = EOQ
        self.H = None
        self.D = None
        self.z = None
        
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
        if self.service_level is not None:
            self.z = self.calculate_z_score(self.service_level)
    
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
    
    def set_parameters(self, demand_rate=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, ordering_cost=None, standard_deviation=None, lead_time=None, service_level=None, weeks_per_year=None, EOQ=None):
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
        if standard_deviation is not None:
            self.standard_deviation = standard_deviation
        if lead_time is not None:
            self.lead_time = lead_time
        if service_level is not None:
            self.service_level = service_level
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
    
    def calculate_z_score(self, service_level):
        """
        Calculate the z-score for a given service level.
        
        Returns:
        float: The calculated z-score.
        """
        return norm.ppf(service_level)
    
    def calculate_safety_stock(self):
        if self.standard_deviation is None or self.lead_time is None or self.z is None:
            raise ValueError("Insufficient parameters to calculate safety stock")
        sigma_dLT = self.standard_deviation * math.sqrt(self.lead_time)
        safety_stock = self.z * sigma_dLT
        print(f"Standard Deviation: {self.standard_deviation}")
        print(f"Lead Time: {self.lead_time}")
        print(f"z-score: {self.z}")
        print(f"sigma_dLT: {sigma_dLT}")
        print(f"Calculated Safety Stock: {safety_stock}")
        return self.custom_round(safety_stock)
    
    def calculate_rop(self):
        """
        Calculate the Reorder Point (ROP).
        
        Returns:
        float: The calculated ROP.
        """
        if self.demand_rate is None or self.lead_time is None:
            raise ValueError("Insufficient parameters to calculate ROP")
        safety_stock = self.calculate_safety_stock()
        return self.custom_round(self.demand_rate * self.lead_time + safety_stock)
    
    def annual_holding_cost(self):
        """
        Calculate the annual total cycle inventory holding costs.
        
        Returns:
        float: The annual total cycle inventory holding costs.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        return self.custom_round((self.EOQ / 2) * self.H)
    
    def annual_ordering_cost(self):
        """
        Calculate the annual total ordering costs.
        
        Returns:
        float: The annual total ordering costs.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        return self.custom_round((self.D / self.EOQ) * self.ordering_cost)
    
    def annual_safety_stock_holding_cost(self):
        """
        Calculate the annual total safety stock holding costs.
        
        Returns:
        float: The annual total safety stock holding costs.
        """
        safety_stock = self.calculate_safety_stock()
        return self.custom_round(safety_stock * self.H)
    
    def total_annual_cost(self):
        """
        Calculate the total annual cost.
        
        Returns:
        float: The total annual cost.
        """
        return self.custom_round(self.annual_holding_cost() + self.annual_ordering_cost() + self.annual_safety_stock_holding_cost())
    
    def time_between_orders(self):
        """
        Calculate the time between two orders (TBO).
        
        Returns:
        float: The time between two orders (TBO) in weeks.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        return self.custom_round(self.EOQ / self.demand_rate)
    
    def custom_round(self, value):
        """
        Custom rounding function that rounds up if the decimal part is greater than 0.2,
        otherwise rounds down.
        """
        integer_part = int(value)
        decimal_part = value - integer_part
        if decimal_part > 0.2:
            return integer_part + 1
        else:
            return integer_part
    
    def plot_costs(self):
        """
        Plot the EOQ model costs.
        """
        if self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        
        # Range of order quantities to plot
        Q_range = linspace(1, 2 * self.EOQ, 500)

        # Calculate costs
        holding_costs = (Q_range / 2) * self.H
        ordering_costs = (self.D / Q_range) * self.ordering_cost
        total_costs = holding_costs + ordering_costs

        figure(figsize=(10, 6))
        plot(Q_range, holding_costs, label='Annual Holding Cost (Q/2 * H)')
        plot(Q_range, ordering_costs, label='Annual Ordering Cost (D/Q * S)')
        plot(Q_range, total_costs, label='Total Annual Cost')
        axvline(self.EOQ, color='purple', linestyle='--', label=f'EOQ = {self.EOQ:.2f}')
        axhline(min(total_costs), color='orange', linestyle='--', label=f'Minimum Total Cost = ${min(total_costs):.2f}')
        xlabel('Order Quantity (Q)')
        ylabel('Cost ($)')
        title('EOQ Model: Costs vs. Order Quantity')
        legend()
        grid(True)
        show()
    
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
            safety_stock = self.calculate_safety_stock()
            safety_stock_cost = self.annual_safety_stock_holding_cost()
            total_cost = self.total_annual_cost()
            TBO = self.time_between_orders()
            ROP = self.calculate_rop()
        except ValueError as e:
            print(f"Error: {e}")
            holding_cost = None
            ordering_cost = None
            safety_stock = None
            safety_stock_cost = None
            total_cost = None
            TBO = None
            ROP = None
        
        print("EOQ Calculation Results:")
        print(f"Demand rate (units/week): {self.demand_rate if self.demand_rate is not None else 'None'}")
        print(f"Demand (units/year): {self.D if self.D is not None else 'None'}")
        print(f"Purchase cost (dollars/unit): {self.purchase_cost if self.purchase_cost is not None else 'None'}")
        print(f"Holding cost rate (annual %): {self.holding_cost_rate if self.holding_cost_rate is not None else 'None'}")
        print(f"Ordering cost (dollars/order): {self.ordering_cost if self.ordering_cost is not None else 'None'}")
        print(f"Weeks per year: {self.weeks_per_year if self.weeks_per_year is not None else 'None'}")
        print(f"Economic Order Quantity (EOQ): {self.EOQ} units" if self.EOQ is not None else "EOQ: None")
        print(f"Annual holding cost per unit (dollars/unit/year): {self.H if self.H is not None else 'None'}")
        print(f"The annual total cycle inventory holding costs are: ${holding_cost}" if holding_cost is not None else "Annual total cycle inventory holding costs: None")
        print(f"The annual total ordering costs are: ${ordering_cost}" if ordering_cost is not None else "Annual total ordering costs: None")
        print(f"Safety stock (units): {safety_stock}" if safety_stock is not None else "Safety stock: None")
        print(f"Annual total safety stock holding costs: ${safety_stock_cost}" if safety_stock_cost is not None else "Annual total safety stock holding costs: None")
        print(f"Total annual cost: ${total_cost}" if total_cost is not None else "Total annual cost: None")
        print(f"The time between two orders (TBO) is: {TBO} weeks" if TBO is not None else "Time between two orders (TBO): None")
        print(f"Reorder Point (ROP): {ROP} units" if ROP is not None else "Reorder Point (ROP): None")

# Test the EOQCalculator with the provided example
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

    eoq_calculator = EOQCalculator(
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

    eoq_calculator.print_results()

if __name__ == "__main__":
    main()
