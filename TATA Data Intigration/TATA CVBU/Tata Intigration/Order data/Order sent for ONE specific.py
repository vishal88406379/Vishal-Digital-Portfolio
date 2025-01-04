import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment

def select_input_file():
    global input_file_path
    input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if input_file_path:
        input_label.config(text=f"Selected File: {input_file_path}")

def select_output_directory():
    global output_directory
    output_directory = filedialog.askdirectory()
    if output_directory:
        output_label.config(text=f"Output Folder: {output_directory}")

def transform_and_save_excel():
    try:
        if not input_file_path or not output_directory:
            messagebox.showwarning("Warning", "Please select an input file and output directory.")
            return
        
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
        ws.title = "Sheet1"  # Set the sheet name to "Sheet1"
        
        # Add rows to the worksheet
        for r in dataframe_to_rows(df.rename(columns=column_mapping), index=False, header=True):
            ws.append(r)
        
        # Adjust column widths based on content length
        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                if cell.value:
                    length = len(str(cell.value))
                    max_length = max(max_length, length)
            adjusted_width = (max_length + 2) * 1.2
            ws.column_dimensions[column].width = adjusted_width
        
        # Ensure the header is not bold
        for cell in ws[1]:
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        output_file_path = f"{output_directory}/Formatted_Output_Order_Data.xlsx"
        wb.save(output_file_path)
        messagebox.showinfo("Success", f"File has been saved successfully at {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Setup the GUI
root = tk.Tk()
root.title("Excel Transformer")

# Enhancements for better GUI appearance
root.geometry('500x300')  # Bigger window size
root.config(bg='light grey')  # Background color

# Heading Label
heading_label = tk.Label(root, text="Order Data Sent", font=('Helvetica', 16, 'bold'), bg='light grey')
heading_label.pack(pady=10)  # Add some vertical padding

input_file_path = ''
output_directory = ''

input_label = tk.Label(root, text="No file selected.", bg='light grey')
input_label.pack(pady=10)

output_label = tk.Label(root, text="No output directory selected.", bg='light grey')
output_label.pack(pady=10)

select_input_button = tk.Button(root, text="Select Input File", command=select_input_file, height=2, width=20)
select_input_button.pack(pady=5)

select_output_button = tk.Button(root, text="Select Output Directory", command=select_output_directory, height=2, width=20)
select_output_button.pack(pady=5)

transform_button = tk.Button(root, text="Transform and Save Output", command=transform_and_save_excel, height=2, width=20)
transform_button.pack(pady=10)

root.mainloop()
