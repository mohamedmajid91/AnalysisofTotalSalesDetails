import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import urllib.request
import os
import sys
import threading
import subprocess
import time
from openpyxl import load_workbook

# ✅ Correct GitHub URL (Make sure your file is hosted correctly)
GITHUB_URL = "https://raw.githubusercontent.com/mohamedmajid91/AnalysisofTotalSalesDetails/main/AnalysisofTotalSalesDetails.exe"

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
    """Process the selected Excel file in a background thread."""
    def process():
        global file_path
        if not file_path:
            messagebox.showerror("Error", "Please select an Excel file first!")
            return
        try:
            # ✅ **Faster Excel Loading**
            df = pd.read_excel(file_path, header=6, engine="openpyxl")
            df = df.iloc[:-1]  # Remove last row if totals
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # Drop unnamed columns

            required_columns = ['Brand', 'SR Name', 'Flavour', 'Sales Quantity']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                messagebox.showerror("Error", f"Missing required columns: {missing_columns}")
                return

            df['Brand'] = df['Brand'].replace({'FRIDAY is a merge': 'FRIDAY'})  # Merge Brands

            if 'Free Cases' in df.columns:
                df['Total Sales'] = df['Sales Quantity'] + df['Free Cases']
            else:
                df['Total Sales'] = df['Sales Quantity']

            # ✅ **Optimized Pivot Tables**
            sr_name_by_brand_and_flavour = df.pivot_table(
                values='Sales Quantity', 
                index=['SR Name'], 
                columns=['Brand', 'Flavour'], 
                aggfunc='sum', 
                fill_value=0
            ).reset_index()

            sr_name_by_brand_and_flavour.columns = [
                ' - '.join(col).strip() if isinstance(col, tuple) else col 
                for col in sr_name_by_brand_and_flavour.columns
            ]

            sr_name_by_brand = df.pivot_table(
                values='Total Sales', 
                index='SR Name', 
                columns='Brand', 
                aggfunc='sum', 
                fill_value=0
            ).reset_index()

            customer_by_brand = df.pivot_table(
                values='Total Sales', 
                index='Customer Name', 
                columns='Brand', 
                aggfunc='sum', 
                fill_value=0
            ).reset_index()

            customer_by_flavour = df.pivot_table(
                values='Total Sales', 
                index='Customer Name', 
                columns='Flavour', 
                aggfunc='sum', 
                fill_value=0
            ).reset_index()

            sr_name_by_total = df.groupby('SR Name').agg({
                'Sales Quantity': 'sum',
                'Free Cases': 'sum' if 'Free Cases' in df.columns else 'sum'
            }).reset_index()
            sr_name_by_total['Total'] = sr_name_by_total['Sales Quantity'] + sr_name_by_total['Free Cases']

            brand_by_total = df.groupby('Brand').agg({
                'Sales Quantity': 'sum',
                'Free Cases': 'sum' if 'Free Cases' in df.columns else 'sum'
            }).reset_index()
            brand_by_total['Total'] = brand_by_total['Sales Quantity'] + brand_by_total['Free Cases']

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

    threading.Thread(target=process, daemon=True).start()

def update_script():
    """Download the latest `.exe` and replace it after restart."""
    def update():
        try:
            script_path = os.path.abspath(sys.argv[0])  # Get current exe path
            temp_path = script_path + ".new"  # Temporary file for update
            backup_path = script_path + ".bak"  # Backup old version

            # ✅ Download the update as a new file
            response = urllib.request.urlopen(GITHUB_URL)
            with open(temp_path, "wb") as temp_file:
                temp_file.write(response.read())

            # ✅ Show update message
            messagebox.showinfo("Update", "✅ Update downloaded! The application will restart.")

            # ✅ Close the application before updating
            root.quit()
            time.sleep(2)  # Wait for the app to fully close

            # ✅ Rename old `.exe` to `.bak` (Backup)
            if os.path.exists(script_path):
                os.rename(script_path, backup_path)

            # ✅ Replace the `.exe` with the new version
            os.rename(temp_path, script_path)

            # ✅ Restart the application after update
            subprocess.Popen([script_path], shell=True)

            sys.exit()  # Exit the old process

        except Exception as e:
            messagebox.showerror("Update Error", f"⚠ Update failed: {e}")

    threading.Thread(target=update, daemon=True).start()

# ✅ **GUI Optimization**
root = tk.Tk()
root.title("Analysis of Total Sales Details")
root.geometry("600x300")

menu_bar = tk.Menu(root)

file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Exit", command=root.quit)
menu_bar.add_cascade(label="File", menu=file_menu)

edit_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Edit", menu=edit_menu)

help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="About", command=lambda: messagebox.showinfo("About", "Excel Data Processor v1.0"))
help_menu.add_command(label="Update", command=update_script)
menu_bar.add_cascade(label="Help", menu=help_menu)

root.config(menu=menu_bar)

button_select = tk.Button(root, text="Select File", width=25, command=select_file)
button_select.pack(pady=10)

label_file = tk.Label(root, text="No file selected", wraplength=580, justify="center")
label_file.pack(pady=5)

button_run = tk.Button(root, text="Execute the Analysis", width=25, command=process_file)
button_run.pack(pady=10)

root.mainloop()
