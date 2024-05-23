import tkinter as tk
from tkinter import ttk, messagebox
from eop_processor import EOQProcessor
import os

class EOQApp:
    def __init__(self, parent):
        self.parent = parent
        self.create_eoq_widgets(parent)
        self.configure_grid_weights(parent)

    def create_eoq_widgets(self, parent):
        # Adding a title
        title_label = ttk.Label(parent, text="EOQ Calculator", font=("Helvetica", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Adding instructions
        instructions_label = ttk.Label(parent, text="Please enter the following parameters. You can leave any field as 'None'.")
        instructions_label.grid(row=1, column=0, columnspan=2, pady=5)

        # Frame for input fields
        input_frame = ttk.Frame(parent)
        input_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        # Labels and dropdown menus for inputs
        self.entries = {}
        parameters = [
            ("Demand Rate (units/week)", "demand_rate"),
            ("Demand Yearly (units/year)", "demand_yearly"),
            ("Purchase Cost (dollars/unit)", "purchase_cost"),
            ("Holding Cost Rate (annual %)", "holding_cost_rate"),
            ("Ordering Cost (dollars/order)", "ordering_cost"),
            ("Standard Deviation (units/week)", "standard_deviation"),
            ("Lead Time (weeks)", "lead_time"),
            ("Service Level (decimal)", "service_level"),
            ("Weeks per Year", "weeks_per_year"),
            ("EOQ", "EOQ")
        ]

        for idx, (label_text, var_name) in enumerate(parameters):
            label = ttk.Label(input_frame, text=label_text)
            label.grid(row=idx, column=0, padx=5, pady=5, sticky=tk.W)
            
            var = tk.StringVar(value="None")
            dropdown = ttk.Combobox(input_frame, textvariable=var, values=["None"])
            dropdown.grid(row=idx, column=1, padx=5, pady=5, sticky="ew")
            self.entries[var_name] = dropdown

        # Frame for buttons
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10, sticky="ew")

        # Calculate button
        self.calculate_button = ttk.Button(button_frame, text="Calculate", command=self.calculate)
        self.calculate_button.grid(row=0, column=0, padx=10)

        # Visualize button
        self.plot_button = ttk.Button(button_frame, text="Visualize", command=self.visualize)
        self.plot_button.grid(row=0, column=1, padx=10)

        # Text area for displaying results
        self.results_text = tk.Text(parent, height=15, width=80, wrap=tk.WORD)
        self.results_text.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        
        # Path for saved Excel file
        self.excel_path = os.path.join(os.path.expanduser("~"), "Downloads", "eoq_results.xlsx")

    def configure_grid_weights(self, parent):
        parent.grid_rowconfigure(0, weight=1)
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_rowconfigure(2, weight=3)
        parent.grid_rowconfigure(3, weight=1)
        parent.grid_rowconfigure(4, weight=5)
        parent.grid_columnconfigure(0, weight=1)
        parent.grid_columnconfigure(1, weight=1)

    def get_input_values(self):
        inputs = {}
        for var_name, dropdown in self.entries.items():
            value = dropdown.get()
            if value == "None":
                inputs[var_name] = None
            else:
                try:
                    inputs[var_name] = float(value) if var_name not in ["weeks_per_year", "EOQ"] else int(value)
                except ValueError:
                    inputs[var_name] = None
        return inputs

    def calculate(self):
        inputs = self.get_input_values()
        if inputs is None:
            return

        print("Inputs:", inputs)  # Debugging: Print inputs to console

        try:
            processor = EOQProcessor(
                demand_rate=inputs.get('demand_rate'),
                demand_yearly=inputs.get('demand_yearly'),
                purchase_cost=inputs.get('purchase_cost'),
                holding_cost_rate=inputs.get('holding_cost_rate'),
                ordering_cost=inputs.get('ordering_cost'),
                standard_deviation=inputs.get('standard_deviation'),
                lead_time=inputs.get('lead_time'),
                service_level=inputs.get('service_level'),
                weeks_per_year=inputs.get('weeks_per_year')
            )

            # Generate and display tables
            input_table = processor.generate_input_table()
            results_table = processor.generate_results_table()

            print("Input Table:\n", input_table)  # Debugging: Print input table
            print("Results Table:\n", results_table)  # Debugging: Print results table

            self.display_results(input_table, results_table)

            # Export results to Excel
            processor.export_to_excel(self.excel_path)
            messagebox.showinfo("Download Successful", "Downloaded Solution Excel. Check your Downloads folder.")

        except Exception as e:
            messagebox.showerror("Calculation Error", f"An error occurred during calculation: {e}")

    def display_results(self, input_table, results_table):
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Inputs Table:\n")
        self.results_text.insert(tk.END, input_table.to_string(index=False))
        self.results_text.insert(tk.END, "\n\nResults Table:\n")
        self.results_text.insert(tk.END, results_table.to_string(index=False))

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
