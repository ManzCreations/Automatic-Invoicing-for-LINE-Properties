import tkinter as tk
from tkinter import ttk
import os
import supporting_strat_auto as ssa
from datetime import datetime
from Initialization import install


##########################
# To install .exe file do the following:
# Open a terminal
# Make sure to be in correct directory
# Type: pip install pyinstaller
# Type: pyinstaller --onefile LINE_AUTO.py
# Go into dist Folder and copy the .exe to Desktop

##########################


class InvoiceGenerator:
    def __init__(self, master):
        # Create the main window
        self.master = master
        master.geometry("400x400")
        master.title("Invoice Generator")
        self.label = tk.Label(master, text="Invoice Generator", font=("Verdana", 12))
        self.label.pack(pady=10, padx=10)

        # Create the 'Install Packages' button
        self.install_packages_button = tk.Button(master, text="Install Packages", command=self.install_packages)
        self.install_packages_button.pack(pady=5)

        # Create the 'Create Invoices' button
        self.button = tk.Button(master, text="Create Invoices", command=self.create_invoices)
        self.button.pack(pady=20)

        # Create the label for the progress bar
        self.command_label = tk.Label(master, text="Progress Bar")
        self.command_label.pack()

        # Create the progress bar
        self.progress_bar = ttk.Progressbar(master, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.pack(pady=10)

        # Create the label for the command window
        self.command_label = tk.Label(master, text="Command Window")
        self.command_label.pack()

        # Create the command window
        self.command_window = tk.Text(master, height=10)
        self.command_window.pack()

    def install_packages(self):
        install(self, 'pandas')
        install(self, 'xlsxwriter')
        install(self, 'openpyxl')

    def run_auto_code(self):
        # Code for generating invoices goes here
        self.log("Creating invoices...")

        path, month = ssa.line_invoice_generation(self)
        self.log("Invoice generation complete.")

        return path, month

    def create_invoices(self):
        self.progress_bar["value"] = 0
        self.progress_bar.update()
        output_dir, output_month = self.run_auto_code()
        os.startfile(output_dir)
        invoice_file = os.path.join(output_dir, "Aviad_BNB_" + output_month + ".xlsx")
        os.startfile(invoice_file)

    def log(self, message, no_time=False):
        # Append the message to the command window

        now = datetime.now()
        current_time = now.strftime("%H:%M:%S")
        if no_time:
            self.command_window.insert(tk.END, message + "\n")
        else:
            self.command_window.insert(tk.END, current_time + ": " + message + "\n")
        self.command_window.update()


root = tk.Tk()
invoice_generator = InvoiceGenerator(root)
root.mainloop()
