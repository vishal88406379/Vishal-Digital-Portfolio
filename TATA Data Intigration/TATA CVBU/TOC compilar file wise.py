import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def process_file(input_path, output_path, progress_var, progress_bar, total_files):
    try:
        # Read the input Excel file
        df = pd.read_excel(input_path)

        # Check if the required columns are present
        if 'SKUCode' not in df.columns or 'Norm' not in df.columns:
            raise ValueError("Input file must contain 'SKUCode' and 'Norm' columns.")

        # Select only the columns of interest and rename them
        df = df[['SKUCode', 'Norm']]
        df.rename(columns={'SKUCode': 'Partnumber', 'Norm': 'Qty'}, inplace=True)

        # Convert 'Partnumber' column to string to ensure text format
        df['Partnumber'] = df['Partnumber'].astype(str)

        # Create the output filename by appending 'Toc'
        output_file_name = os.path.join(output_path, 'Toc_' + os.path.basename(input_path))

        # Create a new workbook and set the sheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"  # Set the sheet name to 'Sheet1'

        # Write dataframe to sheet
        for r_index, r in enumerate(dataframe_to_rows(df, index=False, header=True)):
            ws.append(r)
            # Apply text formatting to 'Partnumber' column
            if r_index > 0:  # Skip the header row
                for cell in ws[r_index + 1]:
                    if cell.column == 1:  # Assuming 'Partnumber' is in column 1 (A)
                        cell.number_format = '@'  # Format as text

        # Save the workbook
        wb.save(output_file_name)
        
        # Update progress bar
        progress_var.set((processed_files[0] / total_files) * 100)
        progress_bar.update()
        
        return f"Processed file: {output_file_name}"

    except Exception as e:
        return f"Error processing {input_path}: {str(e)}"

def browse_input():
    filenames = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    input_file_entry.delete(0, tk.END)
    input_file_entry.insert(0, ';'.join(filenames))

def browse_output():
    foldername = filedialog.askdirectory()
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, foldername)

def run_process():
    input_paths = input_file_entry.get().split(';')
    output_path = output_folder_entry.get()
    
    if not input_paths or not output_path:
        messagebox.showwarning("Input Error", "Please select both input files and output folder.")
        return

    total_files = len(input_paths)
    if total_files == 0:
        messagebox.showwarning("Input Error", "No files selected.")
        return

    # Initialize progress bar
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate", variable=progress_var)
    progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=10)
    progress_var.set(0)

    # Initialize processed files counter
    processed_files[0] = 0
    
    # Process each file
    results = []
    for file in input_paths:
        result = process_file(file, output_path, progress_var, progress_bar, total_files)
        results.append(result)
        processed_files[0] += 1

    # Show completion message
    messagebox.showinfo("Completed", "\n".join(results))

# Set up the GUI
root = tk.Tk()
root.title("Excel File Processor")

tk.Label(root, text="Input Files:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
input_file_entry = tk.Entry(root, width=50)
input_file_entry.grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_input).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_output).grid(row=1, column=2, padx=10, pady=5)

tk.Button(root, text="Run", command=run_process).grid(row=2, column=1, padx=10, pady=20)

# Initialize processed files counter
processed_files = [0]

root.mainloop()
