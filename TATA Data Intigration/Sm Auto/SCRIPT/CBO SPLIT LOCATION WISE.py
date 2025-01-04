import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from openpyxl.styles import Font
from datetime import datetime
import os

def load_input_file():
    input_file_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")]))

def load_location_file():
    location_file_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")]))

def load_output_directory():
    output_directory.set(filedialog.askdirectory())

def generate_unique_filename(base_name, extension, output_dir):
    counter = 1
    new_name = os.path.join(output_dir, f"{base_name}{extension}")
    while os.path.exists(new_name):
        counter += 1
        new_name = os.path.join(output_dir, f"{base_name} {counter}{extension}")
    return new_name

def process_files():
    try:
        input_path = input_file_path.get()
        base_name = "Customer BackOrder SM Auto All Location"
        extension = ".xlsx"
        output_dir = output_directory.get()
        location_path = location_file_path.get()

        if not input_path or not location_path or not output_dir:
            messagebox.showwarning("Warning", "Please select all files and the output directory.")
            return

        df = pd.read_excel(input_path)
        print("Columns in input file:", df.columns.tolist())

        df.rename(columns={'Unnamed: 1': 'PartyName', 'Unnamed: 6': 'Description'}, inplace=True)

        # Ensure 'Product' is treated as text
        if 'Product' in df.columns:
            df['Product'] = df['Product'].astype(str)
        else:
            raise KeyError("Column 'Product' not found.")

        # Check 'ZShip From' column
        if 'ZShip From' not in df.columns:
            raise KeyError("'ZShip From' column not found in input file.")

        # Ensure 'Created On' is a datetime
        if 'Created On' in df.columns:
            df['Created On'] = pd.to_datetime(df['Created On'], dayfirst=True, errors='coerce')
            df['Days'] = (datetime.today() - df['Created On']).dt.days
            cutoff_date = datetime(2024, 6, 26)
            df = df[df['Created On'] > cutoff_date]
        else:
            raise KeyError("Column 'Created On' not found.")

        # Ensure PartyName is string and handle NaN safely
        df['PartyName'] = df['PartyName'].astype(str).fillna('')
        df = df[~df['PartyName'].str.startswith('CSM') & ~df['PartyName'].isin(['Tata Motors Ltd.'])]

        # Ensure External Reference is string and handle NaN safely
        df['External Reference'] = df['External Reference'].astype(str).fillna('')
        df = df[~df['External Reference'].str.strip().str.endswith(('(E)', '(V', '(V)'))]

        # Process numeric columns
        for col in ['Requested Quantity', 'Confirmed Quantity', 'Fulfilled Quantity', 'Invoiced Quantity', 'Pending Qty']:
            if col in df.columns:
                df[col] = df[col].str.replace('ea', '').str.replace(',', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        # Load location mapping file
        try:
            location_mapping_df = pd.read_excel(location_path, sheet_name='Monthly data locations')
        except ValueError:
            messagebox.showerror("Error", "Sheet 'Monthly data locations' not found.")
            return

        if 'ZShip From' not in location_mapping_df.columns or 'Location Name' not in location_mapping_df.columns:
            raise KeyError("'ZShip From' or 'Location Name' column not found in location mapping file.")

        location_mapping = dict(zip(location_mapping_df['ZShip From'], location_mapping_df['Location Name']))
        df['Location'] = df['ZShip From'].map(location_mapping)

        main_output_columns = [
            'Account', 'PartyName', 'Sales Order', 'External Partner',
            'Created On', 'Days', 'Product', 'Description', 'Location',
            'ZShip From', 'External Reference', 'Requested Quantity',
            'Confirmed Quantity', 'Fulfilled Quantity', 'Invoiced Quantity',
            'Pending Qty'
        ]

        for col in main_output_columns:
            if col not in df.columns:
                df[col] = pd.NA

        df = df[main_output_columns]

        # Save main output
        main_output_path = generate_unique_filename(base_name, extension, output_dir)
        df.to_excel(main_output_path, index=False, sheet_name='Sheet1')

        # Retail and TASS processing
        retail_locations = [
            'Varanasi Retail', 'Lucknow STU Retail', 'Lucknow Retail',
            'Gorakhpur retail', 'Allahabad Retail', 'Chopan Retail',
            'Faizabad Retail', 'Fatehpur Retail'
        ]

        tass_locations = [
            'Varanasi TASS', 'LucknowES TASS'
        ]

        # Filter DataFrames for Retail and TASS
        retail_df = df[df['Location'].isin(retail_locations)]
        tass_df = df[df['Location'].isin(tass_locations)]

        retail_output_columns = {
            'Location': 'Location',
            'Sales Order': 'OrderNumber',
            'Created On': 'OrderDate',
            'Account': 'PartyCode',
            'PartyName': 'PartyName',
            'Product': 'PartNumber',
            'Pending Qty': 'Qty'
        }

        # Save Retail output
        if not retail_df.empty:
            retail_output = retail_df.rename(columns=retail_output_columns)[list(retail_output_columns.values())]
            retail_output_path = generate_unique_filename(f"{base_name} Retail", extension, output_dir)
            retail_output.to_excel(retail_output_path, index=False, sheet_name='Sheet1')

        # Save TASS output
        if not tass_df.empty:
            tass_output = tass_df.rename(columns=retail_output_columns)[list(retail_output_columns.values())]
            tass_output_path = generate_unique_filename(f"{base_name} Tass", extension, output_dir)
            tass_output.to_excel(tass_output_path, index=False, sheet_name='Sheet1')

        messagebox.showinfo("Success", f"Files processed and saved successfully:\n- '{main_output_path}'")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        print("Error details:", str(e))

def create_gui():
    global input_file_path, location_file_path, output_directory

    root = tk.Tk()
    root.title("Customer BackOrder SM Auto All Location")

    input_file_path = tk.StringVar()
    location_file_path = tk.StringVar()
    output_directory = tk.StringVar()

    tk.Label(root, text="Select Input Excel File:").pack(pady=5)
    tk.Entry(root, textvariable=input_file_path, width=50).pack(pady=5)
    tk.Button(root, text="Browse", command=load_input_file).pack(pady=5)

    tk.Label(root, text="Select Location Mapping File:").pack(pady=5)
    tk.Entry(root, textvariable=location_file_path, width=50).pack(pady=5)
    tk.Button(root, text="Browse", command=load_location_file).pack(pady=5)

    tk.Label(root, text="Select Output Directory:").pack(pady=5)
    tk.Entry(root, textvariable=output_directory, width=50).pack(pady=5)
    tk.Button(root, text="Browse", command=load_output_directory).pack(pady=5)

    tk.Button(root, text="Run", command=process_files).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
