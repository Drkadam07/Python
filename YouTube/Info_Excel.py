# // Information stored in Excel file

#// by using form to fill Data in Excel file 

import openpyxl
from openpyxl import Workbook
import tkinter as tk
from tkinter import messagebox

# Function to save data to Excel
def save_to_excel(name, address):
    # Load or create the workbook
    try:
        workbook = openpyxl.load_workbook('contacts.xlsx')
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        # Create header if file doesn't exist
        sheet.append(["Name", "Address"])

    # Append the data
    sheet.append([name, address])

    # Save the workbook
    workbook.save('contacts.xlsx')
    messagebox.showinfo("Info", "Data saved successfully!")

# Function to handle the form submission
def submit():
    name = name_entry.get()
    address = address_entry.get()
    if name and address:
        save_to_excel(name, address)
        name_entry.delete(0, tk.END)
        address_entry.delete(0, tk.END)
    else:
        messagebox.showwarning("Warning", "Please fill in all fields")

# Create the GUI window
root = tk.Tk()
root.title("Contact Form")

tk.Label(root, text="Name").grid(row=0, column=0, padx=10, pady=10)
name_entry = tk.Entry(root)
name_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="Address").grid(row=1, column=0, padx=10, pady=10)
address_entry = tk.Entry(root)
address_entry.grid(row=1, column=1, padx=10, pady=10)

submit_button = tk.Button(root, text="Submit", command=submit)
submit_button.grid(row=2, columnspan=2, pady=10)

# Run the GUI event loop
root.mainloop()
