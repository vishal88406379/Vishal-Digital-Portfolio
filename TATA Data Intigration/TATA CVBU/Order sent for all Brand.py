import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
import os
import logging

# Configure logging
logging.basicConfig(filename='app.log', level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

def select_input_files():
    global input_file_paths
    input_file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if input_file_paths:
        file_names = ", ".join([os.path.basename(f) for f in input_file_paths])
        input_label.config(text=f"Selected Files: {file_names}")

def select_output_directory():
    global output_directory
    output_directory = filedialog.askdirectory()
    if output_directory:
        output_label.config(text=f"Output Folder: {output_directory}")

def process_file(input_file_path):
    try:
        column_mapping = {
            'z12-ScsOrderNo': 'SCSOrderNo',
            'z11-Location': 'Location',
            'z10-PartNumber': 'LatestPartNumber',
            'z09-OrderPartNumber': 'OrderPartNumber',
            'z08-Description': 'Description',
            'z07-Rate': 'Rate',
            'z06-SystemMax': 'MaxQty',
            'z05-OpeningStock': 'OpeningStock',
            'z04-OOQ': 'OOQ',
            'z03-CBOQty': 'CustomerBackOrder',
            'z02-SuggestedOrderQty': 'FinalOrderQty',
            'z01-SuggestedOrderValue': 'FinalOrderValue',
            'z00-Avg3MSale': 'Avg3mSale',
            'z000-Category': 'Category'
        }
        use_cols = column_mapping.keys()
        df = pd.read_excel(input_file_path, sheet_name='Order', usecols=lambda x: x in use_cols)
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        
        for r in dataframe_to_rows(df.rename(columns=column_mapping), index=False, header=True):
            ws.append(r)
        
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value))
                    max_length = max(max_length, length)
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
        
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        base_name = os.path.basename(input_file_path).split('.')[0]
        output_file_path = os.path.join(output_directory, f"Order_{base_name}.xlsx")
        
        wb.save(output_file_path)
        return output_file_path
    except Exception as e:
        logging.error(f"Error processing file {input_file_path}: {str(e)}")
        return None

def transform_and_save_excel():
    if not input_file_paths or not output_directory:
        messagebox.showwarning("Warning", "Please select input files and output directory.")
        return

    progress_label.config(text="Processing files...")
    progress_bar['value'] = 0
    progress_bar['maximum'] = len(input_file_paths)
    root.update_idletasks()

    success_files = []
    failed_files = []

    for input_file_path in input_file_paths:
        result = process_file(input_file_path)
        if result:
            success_files.append(result)
        else:
            failed_files.append(input_file_path)
        progress_bar['value'] += 1
        root.update_idletasks()

    if success_files:
        messagebox.showinfo("Success", f"Files saved successfully:\n" + "\n".join(success_files))
    if failed_files:
        messagebox.showerror("Error", f"Failed to process files:\n" + "\n".join(failed_files))

    progress_label.config(text="Processing complete.")

# Setup the GUI
root = tk.Tk()
root.title("Excel Transformer")

# GUI enhancements
root.geometry('600x400')
root.config(bg='light grey')

heading_label = tk.Label(root, text="Order Data Transformer", font=('Helvetica', 16, 'bold'), bg='light grey')
heading_label.pack(pady=10)

input_file_paths = []
output_directory = ''

input_label = tk.Label(root, text="No files selected.", bg='light grey')
input_label.pack(pady=10)

output_label = tk.Label(root, text="No output directory selected.", bg='light grey')
output_label.pack(pady=10)

select_input_button = tk.Button(root, text="Select Input Files", command=select_input_files, height=2, width=20)
select_input_button.pack(pady=5)

select_output_button = tk.Button(root, text="Select Output Directory", command=select_output_directory, height=2, width=20)
select_output_button.pack(pady=5)

transform_button = tk.Button(root, text="Transform and Save Output", command=transform_and_save_excel, height=2, width=20)
transform_button.pack(pady=10)

progress_label = tk.Label(root, text="", bg='light grey')
progress_label.pack(pady=5)

progress_bar = ttk.Progressbar(root, length=400, mode='determinate')
progress_bar.pack(pady=5)

root.mainloop()
