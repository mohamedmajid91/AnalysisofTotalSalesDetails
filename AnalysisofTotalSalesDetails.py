import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import urllib.request
import os
import sys
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# GitHub raw link to fetch updates
GITHUB_URL = "https://raw.githubusercontent.com/mohamedmajid91/AnalysisofTotalSalesDetails/refs/heads/main/AnalysisofTotalSalesDetails.py"

# Global variable to store the file path
file_path = ""

def select_file():
    """Open a file dialog to select an Excel file."""
    global file_path
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
    )
    if file_path:
        label_file.config(text=f"Selected File:\n{file_path}")
    else:
        label_file.config(text="No file selected")

def process_file():
    """Process the selected Excel file and create new sheets with hierarchical headers and aggregations."""
    global file_path
    if not file_path:
        messagebox.showerror("Error", "Please select an Excel file first!")
        return
    
    try:
        # Read the Excel file, setting row 7 (index=6) as the header
        df = pd.read_excel(file_path, header=6)

        # Remove the last row if it's a totals row
        df = df.iloc[:-1]

        # Drop any unnamed columns
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

        # Ensure necessary columns exist
        required_columns = ['Brand', 'SR Name', 'Flavour', 'Sales Quantity']
        for col in required_columns:
            if col not in df.columns:
                messagebox.showerror("Error", f"Missing required column: {col}")
                return

        # Merge "FRIDAY is a merge" into "FRIDAY"
        df['Brand'] = df['Brand'].replace({'FRIDAY is a merge': 'FRIDAY'})

        # Compute Total Sales (if Free Cases exist in the data)
        if 'Free Cases' in df.columns:
            df['Total Sales'] = df['Sales Quantity'] + df['Free Cases']
        else:
            df['Total Sales'] = df['Sales Quantity']

        # 🔹 Pivot Table: SR Name by Brand and Flavour (New Sheet)
        sr_name_by_brand_and_flavour = df.pivot_table(
            values='Sales Quantity', 
            index=['SR Name'], 
            columns=['Brand', 'Flavour'], 
            aggfunc='sum', 
            fill_value=0
        ).reset_index()

        # 🔹 Fix: Flatten MultiIndex Columns
        sr_name_by_brand_and_flavour.columns = [' - '.join(col).strip() if isinstance(col, tuple) else col 
                                                for col in sr_name_by_brand_and_flavour.columns]

        # Pivot Table: SR Name by Brand
        sr_name_by_brand = df.pivot_table(
            values='Total Sales', 
            index='SR Name', 
            columns='Brand', 
            aggfunc='sum', 
            fill_value=0
        ).reset_index()

        # Pivot Table: Customer by Brand
        customer_by_brand = df.pivot_table(
            values='Total Sales', 
            index='Customer Name', 
            columns='Brand', 
            aggfunc='sum', 
            fill_value=0
        ).reset_index()

        # Pivot Table: Customer by Flavour
        customer_by_flavour = df.pivot_table(
            values='Total Sales', 
            index='Customer Name', 
            columns='Flavour', 
            aggfunc='sum', 
            fill_value=0
        ).reset_index()

        # Pivot Table: SR Name by Total (Aggregated Sales & Free Cases)
        sr_name_by_total = df.groupby('SR Name').agg({
            'Sales Quantity': 'sum',
            'Free Cases': 'sum' if 'Free Cases' in df.columns else 'sum'
        }).reset_index()
        sr_name_by_total['Total'] = sr_name_by_total['Sales Quantity'] + sr_name_by_total['Free Cases']

        # Pivot Table: Brand by Total (Aggregated Sales & Free Cases)
        brand_by_total = df.groupby('Brand').agg({
            'Sales Quantity': 'sum',
            'Free Cases': 'sum' if 'Free Cases' in df.columns else 'sum'
        }).reset_index()
        brand_by_total['Total'] = brand_by_total['Sales Quantity'] + brand_by_total['Free Cases']

        # Save to a new Excel file
        output_path = filedialog.asksaveasfilename(
            title="Save Processed File",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx;*.xls"), ("All Files", "*.*")]
        )

        if output_path:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Main Data', index=False)
                sr_name_by_brand_and_flavour.to_excel(writer, sheet_name='SR_Name_by_Brand_and_Flavour', index=False)
                sr_name_by_brand.to_excel(writer, sheet_name='SR_Name_by_Brand', index=False)
                customer_by_brand.to_excel(writer, sheet_name='Customer_by_Brand', index=False)
                customer_by_flavour.to_excel(writer, sheet_name='Customer_by_Flavour', index=False)
                sr_name_by_total.to_excel(writer, sheet_name='SR_Name_by_Total', index=False)
                brand_by_total.to_excel(writer, sheet_name='Brand_by_Total', index=False)

            messagebox.showinfo("Success", f"File processed and saved to:\n{output_path}")
        else:
            messagebox.showwarning("Cancelled", "File save was cancelled.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

def update_script():
    """Download the latest script from GitHub and replace the local file."""
    try:
        script_path = os.path.abspath(sys.argv[0])
        response = urllib.request.urlopen(GITHUB_URL)
        script_content = response.read().decode("utf-8")

        with open(script_path, "w", encoding="utf-8") as file:
            file.write(script_content)

        messagebox.showinfo("Update", "✅ Update completed! Restarting...")
        os.execv(sys.executable, ['python'] + sys.argv)  # Restart script
    except Exception as e:
        messagebox.showerror("Update Error", f"⚠ Update failed: {e}")

# Create the GUI
root = tk.Tk()
root.title("Analysis of Total Sales Details")
root.geometry("600x300")

# Create Menu Bar
menu_bar = tk.Menu(root)

# File Menu
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Exit", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)

# Edit Menu (Can be expanded later)
edit_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Edit", menu=edit_menu)

# Help Menu
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "Excel Data Processor v1.0"))
help_menu.add_command(label="Update", command=update_script)
menu_bar.add_cascade(label="Help", menu=help_menu)

# Configure menu bar in Tkinter window
root.config(menu=menu_bar)

# File Selection Button
button_select = tk.Button(root, text="Select File", width=25, command=select_file)
button_select.pack(pady=10)

label_file = tk.Label(root, text="No file selected", wraplength=580, justify="center")
label_file.pack(pady=5)

# Run Processing Button
button_run = tk.Button(root, text="Execute the Analysis", width=25, command=process_file)
button_run.pack(pady=10)

# Start the Tkinter Event Loop
root.mainloop()
