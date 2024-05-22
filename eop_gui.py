import tkinter as tk
from tkinter import ttk, messagebox
from eop_processor import EOQProcessor
import os

class EOQApp:
    def __init__(self, parent):
        self.parent = parent
        self.create_eoq_widgets(parent)

    def create_eoq_widgets(self, parent):
        # Labels and dropdown menus for inputs
        self.entries = {}
        parameters = [
            ("Demand Rate (units/week)", "demand_rate"),
            ("Demand Yearly (units/year)", "demand_yearly"),
            ("Purchase Cost (dollars/unit)", "purchase_cost"),
            ("Holding Cost Rate (annual %)", "holding_cost_rate"),
            ("Ordering Cost (dollars/order)", "ordering_cost"),
            ("Weeks per Year", "weeks_per_year"),
            ("EOQ", "EOQ")
        ]

        for idx, (label_text, var_name) in enumerate(parameters):
            label = ttk.Label(parent, text=label_text)
            label.grid(row=idx, column=0, padx=10, pady=5, sticky=tk.W)
            
            var = tk.StringVar(value="None")
            dropdown = ttk.Combobox(parent, textvariable=var, values=["None"])
            dropdown.grid(row=idx, column=1, padx=10, pady=5, sticky=tk.W)
            self.entries[var_name] = dropdown

        # Buttons
        self.calculate_button = ttk.Button(parent, text="Calculate", command=self.calculate)
        self.calculate_button.grid(row=len(parameters), column=0, padx=10, pady=10, sticky=tk.W)

        self.plot_button = ttk.Button(parent, text="Visualize", command=self.visualize)
        self.plot_button.grid(row=len(parameters), column=1, padx=10, pady=10, sticky=tk.W)

        # Text area for displaying results
        self.results_text = tk.Text(parent, height=10, width=80)
        self.results_text.grid(row=len(parameters) + 1, column=0, columnspan=3, padx=10, pady=10, sticky=tk.W)
        
        # Path for saved Excel file
        self.excel_path = os.path.join(os.path.expanduser("~"), "Downloads", "eoq_results.xlsx")

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
        processor = EOQProcessor(**inputs)
        input_table = processor.generate_input_table()
        results_table = processor.generate_results_table()
        self.display_results(input_table, results_table)
        processor.export_to_excel(self.excel_path)
        messagebox.showinfo("Download Successful", "Downloaded Solution Excel. Check your Downloads folder.")

    def display_results(self, input_table, results_table):
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, "Inputs Table:\n")
        self.results_text.insert(tk.END, input_table.to_string(index=False))
        self.results_text.insert(tk.END, "\n\nResults Table:\n")
        self.results_text.insert(tk.END, results_table.to_string(index=False))

    def visualize(self):
        inputs = self.get_input_values()
        processor = EOQProcessor(**inputs)
        processor.plot_costs()
