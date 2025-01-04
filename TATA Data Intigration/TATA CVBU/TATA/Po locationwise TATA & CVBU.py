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
        messagebox.showerror("PO Upload Data - Mapping File Error", f"Error loading location mapping: {e}")
        return {}

def process_files(input_folder, output_folder, start_date, end_date, date_format, mapping_file, exclude_v_code):
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Convert start_date and end_date strings to datetime based on selected format
    try:
        start_date = pd.to_datetime(start_date, format=date_format, errors='coerce')
        if pd.isna(start_date):
            raise ValueError("Invalid start date format.")
        end_date = pd.to_datetime(end_date, format=date_format, errors='coerce') if end_date else datetime.today()
    except ValueError as ve:
        messagebox.showerror("PO Upload Data - Date Error", str(ve))
        return

    patterns_to_remove = ['SAP-200', 'SAP-000']
    
    # Load location mapping if provided
    location_mapping = {}
    if mapping_file:
        location_mapping = load_location_mapping(mapping_file)
    
    all_filtered_df = pd.DataFrame()

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            
            try:
                df = pd.read_excel(file_path)
                print("Columns in file:", df.columns.tolist())  # Debug print
            except Exception as e:
                messagebox.showerror("PO Upload Data - File Error", f"Error reading {file_path}: {e}")
                continue

            df['Purchase_Order_Date'] = pd.to_datetime(df['Purchase_Order_Date'], errors='coerce')

            # Apply filtering based on Payer Code if checkbox is selected
            if exclude_v_code and 'Payer Code' in df.columns:
                df = df[~df['Payer Code'].str.startswith('V', na=False)]

            filtered_df = df[(df['Purchase_Order_Date'] >= start_date) & (df['Purchase_Order_Date'] <= end_date)]
            filtered_df = filtered_df[~filtered_df['Order #'].str.contains('|'.join(patterns_to_remove), na=False)]

            if all(col in filtered_df.columns for col in ['Part #', 'Recd Qty', 'Division Name']):
                filtered_df = filtered_df[['Part #', 'Recd Qty', 'Division Name']]
                filtered_df.rename(columns={
                    'Part #': 'Partnumber',
                    'Recd Qty': 'Qty',
                    'Division Name': 'Location'
                }, inplace=True)
            else:
                messagebox.showwarning("PO Upload Data - Column Error", "One or more required columns are missing in the file.")
                continue

            filtered_df['Partnumber'] = filtered_df['Partnumber'].astype(str)

            # Replace location codes with final location names if mapping is provided
            if location_mapping:
                filtered_df['Location'] = filtered_df['Location'].map(location_mapping).fillna(filtered_df['Location'])
            
            all_filtered_df = pd.concat([all_filtered_df, filtered_df], ignore_index=True)

    if all_filtered_df.empty:
        messagebox.showwarning("PO Upload Data - No Data", "No data to process.")
        return

    unique_locations = all_filtered_df['Location'].unique()
    
    for location in unique_locations:
        location_df = all_filtered_df[all_filtered_df['Location'] == location]
        location_file_name = f"{location.replace('/', '_').replace('\\', '_')}.xlsx"
        location_file_path = os.path.join(output_folder, location_file_name)

        with pd.ExcelWriter(location_file_path, engine='xlsxwriter') as writer:
            location_df.to_excel(writer, sheet_name='Sheet1', index=False)
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']
            text_format = workbook.add_format({'num_format': '@'})
            partnumber_col = location_df.columns.get_loc('Partnumber')
            worksheet.set_column(partnumber_col, partnumber_col, None, text_format)
            
            for col_num, value in enumerate(location_df.columns.values):
                max_length = max(location_df[value].astype(str).map(len).max(), len(value)) + 2
                col_letter = chr(65 + col_num)
                worksheet.set_column(f'{col_letter}:{col_letter}', max_length)
            
            worksheet.conditional_format(0, 0, len(location_df) + 1, len(location_df.columns) - 1,
                                         {'type': 'no_blanks',
                                          'format': workbook.add_format({'border': 1})})

    messagebox.showinfo("PO Upload Data - Success", "Files for each location processed and saved successfully.")

def choose_folder(entry_widget):
    folder_path = filedialog.askdirectory(title="Select Input Folder")
    if folder_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder_path)

def choose_file(entry_widget):
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Save Output File As"
    )
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def choose_mapping_file(entry_widget):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx")],
        title="Select Location Mapping File"
    )
    if file_path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, file_path)

def run_process():
    input_folder = input_folder_entry.get()
    output_folder = output_folder_entry.get()
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()
    date_format = date_format_var.get()
    mapping_file = mapping_file_entry.get()
    exclude_v_code = exclude_v_code_var.get()
    
    print(f"Input Folder: {input_folder}")
    print(f"Output Folder: {output_folder}")
    print(f"Start Date: {start_date}")
    print(f"End Date: {end_date}")
    print(f"Date Format: {date_format}")
    print(f"Mapping File: {mapping_file}")
    print(f"Exclude Payer Code Starting with 'V': {exclude_v_code}")
    
    if not input_folder or not output_folder or not start_date:
        messagebox.showwarning("PO Upload Data - Input Error", "Please provide the input folder, output folder, and start date.")
        return
    process_files(input_folder, output_folder, start_date, end_date, date_format, mapping_file, exclude_v_code)

# GUI Setup
root = tk.Tk()
root.title("PO Upload Data - File Processing Tool")

tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=10, pady=10, sticky='e')
input_folder_entry = tk.Entry(root, width=50)
input_folder_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=lambda: choose_folder(input_folder_entry)).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=10, sticky='e')
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=lambda: choose_folder(output_folder_entry)).grid(row=1, column=2, padx=10, pady=10)

tk.Label(root, text="Start Date:").grid(row=2, column=0, padx=10, pady=10, sticky='e')
start_date_entry = tk.Entry(root, width=50)
start_date_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="End Date (optional):").grid(row=3, column=0, padx=10, pady=10, sticky='e')
end_date_entry = tk.Entry(root, width=50)
end_date_entry.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="Date Format:").grid(row=4, column=0, padx=10, pady=10, sticky='e')
date_format_var = tk.StringVar(value='%d/%m/%Y')  # Default format
date_format_menu = tk.OptionMenu(root, date_format_var, '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d')
date_format_menu.grid(row=4, column=1, padx=10, pady=10, sticky='w')

tk.Label(root, text="Output File:").grid(row=5, column=0, padx=10, pady=10, sticky='e')
output_file_entry = tk.Entry(root, width=50)
output_file_entry.grid(row=5, column=1, padx=10, pady=10)
tk.Button(root, text="Save As...", command=lambda: choose_file(output_file_entry)).grid(row=5, column=2, padx=10, pady=10)

tk.Label(root, text="Location Mapping File:").grid(row=6, column=0, padx=10, pady=10, sticky='e')
mapping_file_entry = tk.Entry(root, width=50)
mapping_file_entry.grid(row=6, column=1, padx=10, pady=10)
tk.Button(root, text="Browse...", command=lambda: choose_mapping_file(mapping_file_entry)).grid(row=6, column=2, padx=10, pady=10)

exclude_v_code_var = tk.BooleanVar()
tk.Checkbutton(root, text="Exclude Payer Code Starting with 'V'", variable=exclude_v_code_var).grid(row=7, column=1, padx=10, pady=10, sticky='w')

tk.Button(root, text="Run", command=run_process).grid(row=8, column=1, padx=10, pady=20)

root.mainloop()
