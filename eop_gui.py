import tkinter as tk
from tkinter import ttk, messagebox
from eop_processor import EOQProcessor
import os

valid_vals = ["", "None"]
class EOQApp:
    def __init__(self, parent):
        self.parent = parent
        self.create_eoq_widgets(parent)
        self.configure_grid_weights(parent)

    def error_ret(self, new_value):
        messagebox.showerror("Invalid Entry", f"'{new_value}' is not a valid integer.")
        return False

    def validate_val(self, new_value):
        if new_value in valid_vals or new_value.isdigit():
            return True
        else:
            return self.error_ret(new_value)

    def is_valid_decimal(self, value):
        try:
            float(value)
            return True
        except ValueError:
            return False

    def validate_decimal(self, new_value):
        if new_value in valid_vals or self.is_valid_decimal(new_value):
            return True
        else:
            return self.error_ret(new_value)

    def create_eoq_widgets(self, parent):
        # Adding a title
        title_label = ttk.Label(parent, text="EOQ Calculator", font=("Helvetica", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)

        # Adding instructions
        instructions_label = ttk.Label(parent, text="Please enter the following parameters. You can leave any field as 'None'.")
        instructions_label.grid(row=1, column=0, columnspan=3, pady=5)

        # Frame for input fields
        input_frame = ttk.Frame(parent)
        input_frame.grid(row=2, column=0, padx=10, pady=10, sticky="nsew")

        # Labels and dropdown menus for inputs
        self.entries = {}
        parameters = [
            ("Demand Rate (units/day)", "demand_rate", self.parent.register(self.validate_val)),
            ("Demand Rate (units/week)", "demand_weekly", self.parent.register(self.validate_val)),
            ("Demand Yearly (units/year)", "demand_yearly", self.parent.register(self.validate_val)),
            ("Purchase Cost (dollars/unit)", "purchase_cost", self.parent.register(self.validate_decimal)),
            ("Holding Cost Rate (annual %)", "holding_cost_rate", self.parent.register(self.validate_decimal)),
            ("Holding Cost per Unit (dollars/unit/year)", "holding_cost_per_unit", self.parent.register(self.validate_decimal)),
            ("Ordering Cost (dollars/order)", "ordering_cost", self.parent.register(self.validate_decimal)),
            ("Standard Deviation (units/day)", "standard_deviation_per_day", self.parent.register(self.validate_decimal)),
            ("Lead Time (weeks)", "lead_time", self.parent.register(self.validate_val)),
            ("Lead Time (days)", "lead_time_days", self.parent.register(self.validate_val)),
            ("Service Level (decimal)", "service_level", self.parent.register(self.validate_decimal)),
            ("Weeks per Year", "weeks_per_year", self.parent.register(self.validate_val)),
            ("Days per Year", "days_per_year", self.parent.register(self.validate_val)),
            ("EOQ", "EOQ", self.parent.register(self.validate_decimal)),
            ("Toggle Holding Stock (yes/no)", "toggle_holding_stock", None)
        ]

        for idx, (label_text, var_name, validate_command) in enumerate(parameters):
            label = ttk.Label(input_frame, text=label_text)
            label.grid(row=idx, column=0, padx=5, pady=5, sticky=tk.W)
            var = tk.StringVar(value="None")
            if var_name == "toggle_holding_stock":
                dropdown = ttk.Combobox(input_frame, textvariable=var, values=["yes", "no"])
            else:
                if validate_command:
                    dropdown = ttk.Combobox(input_frame, textvariable=var, values=[valid_vals[1]], validate='key', validatecommand=(validate_command, '%P'))
                else:
                    dropdown = ttk.Combobox(input_frame, textvariable=var, values=[valid_vals[1]])
            dropdown.grid(row=idx, column=1, padx=5, pady=5, sticky="ew")
            self.entries[var_name] = dropdown

        # Frame for buttons
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=1, pady=10, sticky="w")

        # Calculate EOQ button
        self.calculate_eoq_button = ttk.Button(button_frame, text="Calculate EOQ Only", command=self.calculate_eoq_only)
        self.calculate_eoq_button.grid(row=0, column=0, padx=10)

        # Calculate Full button
        self.calculate_full_button = ttk.Button(button_frame, text="Calculate Full Set", command=self.calculate_full_set)
        self.calculate_full_button.grid(row=0, column=1, padx=10)

        # Visualize button
        self.plot_button = ttk.Button(button_frame, text="Visualize", command=self.visualize)
        self.plot_button.grid(row=0, column=2, padx=10)

        # Text area for displaying results
        self.results_text = tk.Text(parent, height=30, width=80, wrap=tk.WORD)
        self.results_text.grid(row=2, column=1, rowspan=2, padx=10, pady=10, sticky="nsew")

        # Path for saved Excel file
        self.excel_path = os.path.join(os.path.expanduser("~"), "Downloads", "eoq_results.xlsx")

    def configure_grid_weights(self, parent):
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_rowconfigure(3, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)

    def get_input_values(self):
        inputs = {}
        for var_name, dropdown in self.entries.items():
            value = dropdown.get()
            if value == valid_vals[1]:
                inputs[var_name] = None
            elif var_name == "toggle_holding_stock":
                inputs[var_name] = value.lower() == "yes"
            else:
                try:
                    if var_name in ["weeks_per_year", "days_per_year", "EOQ"]:
                        inputs[var_name] = int(value)
                    else:
                        inputs[var_name] = float(value)
                except ValueError:
                    inputs[var_name] = None
        return inputs

    def calculate_eoq_only(self):
        self.calculate(full_set=False)

    def calculate_full_set(self):
        self.calculate(full_set=True)

    def calculate(self, full_set):
        inputs = self.get_input_values()
        if inputs is None:
            return

        print("Inputs:", inputs)  # Debugging: Print inputs to console

        try:
            processor = EOQProcessor(
                demand_rate=inputs.get('demand_rate'),
                demand_weekly=inputs.get('demand_weekly'),
                demand_yearly=inputs.get('demand_yearly'),
                purchase_cost=inputs.get('purchase_cost'),
                holding_cost_rate=inputs.get('holding_cost_rate'),
                holding_cost_per_unit=inputs.get('holding_cost_per_unit'),
                ordering_cost=inputs.get('ordering_cost'),
                standard_deviation=inputs.get('standard_deviation_per_day'),
                lead_time=inputs.get('lead_time'),
                lead_time_days=inputs.get('lead_time_days'),
                service_level=inputs.get('service_level'),
                weeks_per_year=inputs.get('weeks_per_year'),
                days_per_year=inputs.get('days_per_year'),
                EOQ=inputs.get('EOQ'),
                toggle_holding_stock=inputs.get('toggle_holding_stock')
            )

            if full_set:
                # Generate and display tables
                input_table = processor.generate_input_table()
                results_table = processor.generate_results_table()

                print("Input Table:\n", input_table)  # Debugging: Print input table
                print("Results Table:\n", results_table)  # Debugging: Print results table

                self.display_results(input_table, results_table)

                # Export results to Excel
                processor.export_to_excel(self.excel_path)
                messagebox.showinfo("Download Successful", "Downloaded Solution Excel. Check your Downloads folder.")
            else:
                processor.calculator.print_results(full_set=False)
                self.display_eoq_only_results(processor)

        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred during calculation: {e}")

    def display_results(self, input_table, results_table):
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Inputs Table:\n")
        self.results_text.insert(tk.END, input_table.to_string(index=False))
        self.results_text.insert(tk.END, "\n\nResults Table:\n")
        self.results_text.insert(tk.END, results_table.to_string(index=False))

    def display_eoq_only_results(self, processor):
        input_table = processor.generate_input_table()
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Inputs Table:\n")
        self.results_text.insert(tk.END, input_table.to_string(index=False))
        self.results_text.insert(tk.END, "\n\nEOQ Calculation Results:\n")
        self.results_text.insert(tk.END, f"Economic Order Quantity (EOQ): {processor.calculator.EOQ} units\n")
        self.results_text.insert(tk.END, f"Number of Orders per Year: {processor.calculator.number_of_orders_per_year()}\n")
        self.results_text.insert(tk.END, f"Time Between Orders (TBO): {processor.calculator.time_between_orders()} days\n")

    def visualize(self):
        inputs = self.get_input_values()
        if inputs is None:
            return
        processor = EOQProcessor(**inputs)
        processor.calculator.update_calculations()  # Ensure calculations are performed
        processor.plot_costs()

# Setup and run the application
if __name__ == "__main__":
    root = tk.Tk()
    root.title("EOQ Calculator")
    app = EOQApp(root)
    root.mainloop()
