import math
from numpy import linspace
import matplotlib.pyplot as plt
from scipy.stats import norm

class EOQCalculator:
    def __init__(self, demand_rate=None, demand_weekly=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, holding_cost_per_unit=None, ordering_cost=None, standard_deviation=None, standard_deviation_per_day=None, lead_time=None, lead_time_days=None, service_level=None, weeks_per_year=None, days_per_year=365, EOQ=None, toggle_holding_stock=True):
        self.demand_rate = demand_rate  # units per day
        self.demand_weekly = demand_weekly # units per week
        self.demand_yearly = demand_yearly  # units per year
        self.purchase_cost = purchase_cost  # cost per unit
        self.holding_cost_rate = holding_cost_rate  # holding cost rate per year
        self.holding_cost_per_unit = holding_cost_per_unit  # holding cost per unit per year
        self.ordering_cost = ordering_cost  # cost per order
        self.standard_deviation = standard_deviation  # standard deviation of demand
        self.standard_deviation_per_day = standard_deviation_per_day  # standard deviation of demand per day
        self.lead_time = lead_time  # lead time in weeks
        self.lead_time_days = lead_time_days  # lead time in days
        self.service_level = service_level  # service level as a decimal (e.g., 0.9 for 90%)
        self.weeks_per_year = weeks_per_year if weeks_per_year is not None else days_per_year / 7  # number of weeks in a year
        self.days_per_year = days_per_year  # number of operational days in a year
        self.EOQ = EOQ
        self.H = None
        self.D = None
        self.z = None
        self.toggle_holding_stock = toggle_holding_stock

        self.update_calculations()

    def update_calculations(self):
        if self.days_per_year is None:
            self.days_per_year = self.weeks_per_year * 7
        if self.demand_rate is not None:
            self.D = self.demand_rate * self.days_per_year  # annual demand in units/year
        elif self.demand_weekly is not None:
            self.D = self.demand_weekly * self.weeks_per_year # convert to yearly demand
            self.demand_rate = self.D / self.days_per_year
        elif self.demand_yearly is not None:
            self.D = self.demand_yearly
            self.demand_rate = self.D / self.days_per_year  # calculate daily demand from annual demand
        if self.holding_cost_rate is not None and self.purchase_cost is not None:
            self.H = self.holding_cost_rate * self.purchase_cost  # annual holding cost per unit
        if self.holding_cost_per_unit is not None:
            self.H = self.holding_cost_per_unit
        if self.D is not None and self.ordering_cost is not None and self.H is not None and self.EOQ is None:
            self.EOQ = self.calculate_eoq()
        if self.EOQ is not None and self.EOQ > 0:
            self.solve_missing_parameters()
        if self.service_level is not None:
            self.z = self.calculate_z_score(self.service_level)
        if self.standard_deviation_per_day is not None:
            self.standard_deviation = self.standard_deviation_per_day * math.sqrt(self.days_per_year)
        # if self.lead_time_days is not None and self.days_per_year is not None and self.weeks_per_year is not None:
        #     self.lead_time = self.lead_time_days / (self.days_per_year / self.weeks_per_year)
        # elif self.lead_time is not None and self.days_per_year is not None and self.weeks_per_year is not None:
        #     self.lead_time_days = self.lead_time * (self.days_per_year / self.weeks_per_year)

    def solve_missing_parameters(self):
        if self.D is None and self.ordering_cost is not None and self.H is not None:
            self.D = (self.EOQ ** 2 * self.H) / (2 * self.ordering_cost)
            self.demand_rate = self.D / self.days_per_year
        if self.ordering_cost is None and self.D is not None and self.H is not None:
            self.ordering_cost = (self.EOQ ** 2 * self.H) / (2 * self.D)
        if self.H is None and self.D is not None and self.ordering_cost is not None:
            self.H = (2 * self.D * self.ordering_cost) / (self.EOQ ** 2)
        if self.holding_cost_rate is None and self.H is not None and self.purchase_cost is not None:
            self.holding_cost_rate = self.H / self.purchase_cost
        if self.purchase_cost is None and self.H is not None and self.holding_cost_rate is not None:
            self.purchase_cost = self.H / self.holding_cost_rate
        if self.demand_rate is None and self.D is not None and self.days_per_year is not None:
            self.demand_rate = self.D / self.days_per_year

    def set_parameters(self, demand_rate=None, demand_weekly=None, demand_yearly=None, purchase_cost=None, holding_cost_rate=None, holding_cost_per_unit=None, ordering_cost=None, standard_deviation=None, standard_deviation_per_day=None, lead_time=None, lead_time_days=None, service_level=None, weeks_per_year=None, days_per_year=None, EOQ=None, toggle_holding_stock=None):
        if demand_rate is not None:
            self.demand_rate = demand_rate
        if demand_weekly is not None:
            self.demand_weekly = demand_weekly
        if demand_yearly is not None:
            self.demand_yearly = demand_yearly
        if purchase_cost is not None:
            self.purchase_cost = purchase_cost
        if holding_cost_rate is not None:
            self.holding_cost_rate = holding_cost_rate
        if holding_cost_per_unit is not None:
            self.holding_cost_per_unit = holding_cost_per_unit
        if ordering_cost is not None:
            self.ordering_cost = ordering_cost
        if standard_deviation is not None:
            self.standard_deviation = standard_deviation
        if standard_deviation_per_day is not None:
            self.standard_deviation_per_day = standard_deviation_per_day
        if lead_time is not None:
            self.lead_time = lead_time
        if lead_time_days is not None:
            self.lead_time_days = lead_time_days
        if service_level is not None:
            self.service_level = service_level
        if weeks_per_year is not None:
            self.weeks_per_year = weeks_per_year
        if days_per_year is not None:
            self.days_per_year = days_per_year
        if EOQ is not None:
            self.EOQ = EOQ
        if toggle_holding_stock is not None:
            self.toggle_holding_stock = toggle_holding_stock

        self.update_calculations()

    def calculate_eoq(self):
        if self.D is None or self.ordering_cost is None or self.H is None:
            raise ValueError("Insufficient parameters to calculate EOQ")
        if self.D == 0 or self.H == 0 or self.ordering_cost == 0:
            raise ValueError("D, H, and ordering cost must be non-zero for EOQ calculation")
        print("My Params", self.D, self.ordering_cost, self.H)
        return math.sqrt((2 * self.D * self.ordering_cost) / self.H)

    def calculate_z_score(self, service_level):
        return norm.ppf(service_level)

    def calculate_safety_stock(self):
        if not self.toggle_holding_stock:
            return 0
        if self.standard_deviation is None or self.z is None:
            raise ValueError("Insufficient parameters to calculate safety stock")
        if self.lead_time is not None:
            sigma_dLT = self.standard_deviation * math.sqrt((self.lead_time*self.days_per_year/self.weeks_per_year))
        elif self.lead_time_days is not None:
            sigma_dLT = self.standard_deviation * math.sqrt(self.lead_time_days)
        else:
            raise ValueError("Insufficient parameters to calculate safety stock")
        safety_stock = self.z * sigma_dLT
        return round(safety_stock, 1)

    def calculate_rop(self):
        if not self.toggle_holding_stock:
            return 0
        if self.demand_rate is None:
            raise ValueError("Insufficient parameters to calculate ROP")
        safety_stock = self.calculate_safety_stock()
        if self.lead_time is not None:
            return round(self.demand_rate * self.lead_time * (self.days_per_year / self.weeks_per_year) + safety_stock, 1)
        elif self.lead_time_days is not None:
            return round(self.demand_rate * self.lead_time_days + safety_stock, 1)
        else:
            raise ValueError("Insufficient parameters to calculate ROP")

    def annual_holding_cost(self):
        if self.EOQ is None or self.EOQ == 0:
            raise ValueError("EOQ must be calculated and non-zero before calculating holding cost")
        if self.H is None or self.H == 0:
            raise ValueError("Holding cost per unit (H) must be non-zero before calculating holding cost")
        return round((self.EOQ / 2) * self.H, 1)

    def annual_ordering_cost(self):
        if self.EOQ is None or self.EOQ == 0:
            raise ValueError("EOQ must be calculated and non-zero before calculating ordering cost")
        return round((self.D / self.EOQ) * self.ordering_cost, 1)

    def annual_safety_stock_holding_cost(self):
        if not self.toggle_holding_stock:
            return 0
        safety_stock = self.calculate_safety_stock()
        return round(safety_stock * self.H, 1)

    def total_annual_cost(self):
        holding_cost = self.annual_holding_cost()
        ordering_cost = self.annual_ordering_cost()
        safety_stock_cost = self.annual_safety_stock_holding_cost() if self.toggle_holding_stock else 0
        return round(holding_cost + ordering_cost + safety_stock_cost, 1)

    def time_between_orders(self):
        if self.EOQ is None or self.EOQ == 0:
            raise ValueError("EOQ must be calculated and non-zero before calculating time between orders")
        if self.demand_rate == 0:
            raise ValueError("Demand rate cannot be zero for time between orders calculation")
        return round(self.EOQ / self.demand_rate, 1)

    def number_of_orders_per_year(self):
        if self.EOQ is None or self.EOQ == 0:
            raise ValueError("EOQ must be calculated and non-zero before calculating number of orders")
        return round(self.D / self.EOQ, 1)

    def plot_costs(self):
        if self.EOQ is None or self.EOQ == 0:
            raise ValueError("EOQ must be calculated and non-zero before plotting costs")

        Q_range = linspace(1, 2 * self.EOQ, 500)

        holding_costs = (Q_range / 2) * self.H if self.toggle_holding_stock else 0
        ordering_costs = (self.D / Q_range) * self.ordering_cost
        total_costs = holding_costs + ordering_costs

        plt.figure(figsize=(10, 6))
        plt.plot(Q_range, holding_costs, label='Annual Holding Cost (Q/2 * H)', color='green')
        plt.plot(Q_range, ordering_costs, label='Annual Ordering Cost (D/Q * S)', color='blue')
        plt.plot(Q_range, total_costs, label='Total Annual Cost', color='red')
        plt.axvline(self.EOQ, color='purple', linestyle='--', label=f'EOQ = {self.EOQ:.2f}')
        plt.axhline(min(total_costs), color='orange', linestyle='--', label=f'Minimum Total Cost = ${min(total_costs):.2f}')
        plt.xlabel('Order Quantity (Q)')
        plt.ylabel('Cost ($)')
        plt.title('EOQ Model: Costs vs. Order Quantity')
        plt.legend()
        plt.grid(True)
        plt.show()

    def print_results(self, full_set=True):
        try:
            if self.EOQ is None or self.EOQ == 0:
                self.calculate_eoq()
            holding_cost = self.annual_holding_cost()
            ordering_cost = self.annual_ordering_cost()
            safety_stock = self.calculate_safety_stock() if self.toggle_holding_stock else None
            safety_stock_cost = self.annual_safety_stock_holding_cost() if self.toggle_holding_stock else None
            total_cost = self.total_annual_cost()
            TBO = self.time_between_orders()
            ROP = self.calculate_rop() if self.toggle_holding_stock else None
            orders_per_year = self.number_of_orders_per_year()
        except ValueError as e:
            print(f"Error: {e}")
            holding_cost = None
            ordering_cost = None
            safety_stock = None
            safety_stock_cost = None
            total_cost = None
            TBO = None
            ROP = None
            orders_per_year = None

        print("EOQ Calculation Results:")
        print(f"Demand rate (units/day): {self.demand_rate if self.demand_rate is not None else 'None'}")
        print(f"Demand (units/week): {self.D if self.D is not None else 'None'}")
        print(f"Demand (units/year): {self.D if self.D is not None else 'None'}")
        print(f"Purchase cost (dollars/unit): {self.purchase_cost if self.purchase_cost is not None else 'None'}")
        print(f"Holding cost rate (annual %): {self.holding_cost_rate if self.holding_cost_rate is not None else 'None'}")
        print(f"Holding cost per unit (dollars/unit/year): {self.holding_cost_per_unit if self.holding_cost_per_unit is not None else 'None'}")
        print(f"Ordering cost (dollars/order): {self.ordering_cost if self.ordering_cost is not None else 'None'}")
        print(f"Weeks per year: {self.weeks_per_year if self.weeks_per_year is not None else 'None'}")
        print(f"Days per year: {self.days_per_year if self.days_per_year is not None else 'None'}")
        print(f"Economic Order Quantity (EOQ): {self.EOQ} units" if self.EOQ is not None else "EOQ: None")
        print(f"Annual holding cost per unit (dollars/unit/year): {self.H if self.H is not None else 'None'}")
        print(f"Annual total cycle inventory holding costs: ${holding_cost}" if holding_cost is not None else "Annual total cycle inventory holding costs: None")
        print(f"Annual total ordering costs: ${ordering_cost}" if ordering_cost is not None else "Annual total ordering costs: None")
        if self.toggle_holding_stock:
            print(f"Safety stock (units): {safety_stock}" if safety_stock is not None else "Safety stock: None")
            print(f"Annual total safety stock holding costs: ${safety_stock_cost}" if safety_stock_cost is not None else "Annual total safety stock holding costs: None")
            print(f"Reorder Point (ROP): {ROP} units" if ROP is not None else "Reorder Point (ROP): None")
        print(f"Total annual cost: ${total_cost}" if total_cost is not None else "Total annual cost: None")
        print(f"Time between orders (TBO): {TBO} days" if TBO is not None else "Time between orders (TBO): None")
        print(f"Number of orders per year: {orders_per_year}" if orders_per_year is not None else "Number of orders per year: None")

def main():
    calculator = EOQCalculator(
        demand_rate=15.0,  # units per day
        demand_yearly=None,
        purchase_cost=11.7,  # dollars per unit
        holding_cost_rate=0.28,
        holding_cost_per_unit=None,  # dollars per unit per year
        ordering_cost=54.0,  # dollars per order
        standard_deviation=None,
        standard_deviation_per_day=6.124,  # units per day
        lead_time=None,
        lead_time_days=18.0,  # days
        service_level=0.8,  # 80% service level
        weeks_per_year=None,  # Default value
        days_per_year=312,  # number of days in a year
        EOQ=None,
        toggle_holding_stock=True  # toggle holding stock
    )
    calculator.print_results(full_set=True)

if __name__ == "__main__":
    main()
