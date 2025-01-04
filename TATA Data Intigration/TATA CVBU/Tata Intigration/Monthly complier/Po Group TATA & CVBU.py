import pandas as pd
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import xlsxwriter

def load_location_mapping(mapping_file):
    """Load location mapping from an Excel file."""
    try:
        mapping_df = pd.read_excel(mapping_file)
        print("Columns in mapping file:", mapping_df.columns.tolist())  # Debug print
        return dict(zip(mapping_df['Code'], mapping_df['Final Location']))
    except Exception as e:
        messagebox.showerror("Mapping File Error", f"Error loading location mapping: {e}")
        raise

def process_files(input_folder, output_file, start_date, date_format, mapping_file):
    # Ensure the output folder exists
    output_folder = os.path.dirname(output_file)
    os.makedirs(output_folder, exist_ok=True)

    # Convert start_date string to datetime based on selected format
    try:
        start_date = pd.to_datetime(start_date, format=date_format, errors='coerce')
        if pd.isna(start_date):
            raise ValueError("Invalid date format.")
    except ValueError:
        messagebox.showerror("Date Error", "Invalid date format. Please use the correct format.")
        return

    today = datetime.today()
    patterns_to_remove = ['EICPUR', 'ICPPUR', 'CPPUR', 'ECPUR', 'SAP-200', 'SAP-300', 'SAP-000']
    location_mapping = load_location_mapping(mapping_file)
    all_filtered_df = pd.DataFrame()

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            
            try:
                df = pd.read_excel(file_path)
                print("Columns in file:", df.columns.tolist())  # Debug print
            except Exception as e:
                messagebox.showerror("File Error", f"Error reading {file_path}: {e}")
                continue

            df['Purchase_Order_Date'] = pd.to_datetime(df['Purchase_Order_Date'], errors='coerce')
            filtered_df = df[(df['Purchase_Order_Date'] >= start_date) & (df['Purchase_Order_Date'] <= today)]
            filtered_df = filtered_df[~filtered_df['Order #'].str.contains('|'.join(patterns_to_remove), na=False)]

            if all(col in filtered_df.columns for col in ['Part #', 'Recd Qty', 'Division Name']):
                filtered_df = filtered_df[['Part #', 'Recd Qty', 'Division Name']]
                filtered_df.rename(columns={
                    'Part #': 'Partnumber',
                    'Recd Qty': 'Qty',
                    'Division Name': 'Location'
                }, inplace=True)
            else:
                messagebox.showwarning("Column Error", "One or more required columns are missing in the file.")
                continue

            filtered_df['Partnumber'] = filtered_df['Partnumber'].astype(str)
            filtered_df['Location'] = filtered_df['Location'].map(location_mapping).fillna('n/a')
            all_filtered_df = pd.concat([all_filtered_df, filtered_df], ignore_index=True)

    if not output_file.lower().endswith('.xlsx'):
        output_file += '.xlsx'

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        all_filtered_df.to_excel(writer, sheet_name='FilteredData', index=False)
        workbook = writer.book
        worksheet = writer.sheets['FilteredData']
        text_format = workbook.add_format({'num_format': '@'})
        partnumber_col = all_filtered_df.columns.get_loc('Partnumber')
        worksheet.set_column(partnumber_col, partnumber_col, None, text_format)
        
        for col_num, value in enumerate(all_filtered_df.columns.values):
            max_length = max(all_filtered_df[value].astype(str).map(len).max(), len(value)) + 2
            col_letter = chr(65 + col_num)
            worksheet.set_column(f'{col_letter}:{col_letter}', max_length)
        
        worksheet.conditional_format(0, 0, len(all_filtered_df) + 1, len(all_filtered_df.columns) - 1,
                                     {'type': 'no_blanks',
                                      'format': workbook.add_format({'border': 1})})
        
    messagebox.showinfo("Success", "All files processed and saved successfully with 'FilteredData' sheet.")

def choose_folder():
    folder_path = filedialog.askdirectory(title="Select Input Folder")
    if folder_path:
        input_folder_entry.delete(0, tk.END)
        input_folder_entry.insert(0, folder_path)

def choose_file():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Output File As"
    )
    if file_path:
        output_file_entry.delete(0, tk.END)
        output_file_entry.insert(0, file_path)

def choose_mapping_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")],
        title="Select Location Mapping File"
    )
    if file_path:
        mapping_file_entry.delete(0, tk.END)
        mapping_file_entry.insert(0, file_path)

def run_process():
    input_folder = input_folder_entry.get()
    output_file = output_file_entry.get()
    start_date = start_date_entry.get()
    date_format = date_format_var.get()
    mapping_file = mapping_file_entry.get()
    
    print(f"Input Folder: {input_folder}")
    print(f"Output File: {output_file}")
    print(f"Start Date: {start_date}")
    print(f"Date Format: {date_format}")
    print(f"Mapping File: {mapping_file}")
    
    if not input_folder or not output_file or not start_date or not mapping_file:
        messagebox.showwarning("Input Error", "Please provide all required inputs.")
        return
    process_files(input_folder, output_file, start_date, date_format, mapping_file)

# GUI Setup
root = tk.Tk()
root.title("File Processing Tool")

tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
input_folder_entry = tk.Entry(root, width=50)
input_folder_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=choose_folder).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Output File:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
output_file_entry = tk.Entry(root, width=50)
output_file_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=choose_file).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Start Date:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
start_date_entry = tk.Entry(root, width=50)
start_date_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="Date Format:").grid(row=3, column=0, padx=10, pady=10, sticky='e')
date_format_var = tk.StringVar(value='%d/%m/%Y')  # Default format
date_format_menu = tk.OptionMenu(root, date_format_var, '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d')
date_format_menu.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Location Mapping File:").grid(row=4, column=0, padx=10, pady=10, sticky='e')
mapping_file_entry = tk.Entry(root, width=50)
mapping_file_entry.grid(row=4, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=choose_mapping_file).grid(row=4, column=2, padx=10, pady=10)

tk.Button(root, text="Run Process", command=run_process).grid(row=5, column=1, padx=10, pady=20)

root.mainloop()
