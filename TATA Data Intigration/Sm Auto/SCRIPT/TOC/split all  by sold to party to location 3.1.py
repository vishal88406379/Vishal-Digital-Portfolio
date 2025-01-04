import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

class ExcelSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Splitter")

        # Input file
        self.input_file_label = tk.Label(root, text="Select Excel File:")
        self.input_file_label.pack(pady=5)
        
        self.input_file_entry = tk.Entry(root, width=50)
        self.input_file_entry.pack(pady=5)
        
        self.browse_input_button = tk.Button(root, text="Browse", command=self.browse_input_file)
        self.browse_input_button.pack(pady=5)

        # Output folder
        self.output_folder_label = tk.Label(root, text="Select Output Folder:")
        self.output_folder_label.pack(pady=5)
        
        self.output_folder_entry = tk.Entry(root, width=50)
        self.output_folder_entry.pack(pady=5)
        
        self.browse_output_button = tk.Button(root, text="Browse", command=self.browse_output_folder)
        self.browse_output_button.pack(pady=5)
        
        # Run button
        self.run_button = tk.Button(root, text="Split and Save", command=self.split_and_save)
        self.run_button.pack(pady=20)
        
    def browse_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_file_entry.delete(0, tk.END)
            self.input_file_entry.insert(0, file_path)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder_path)

    def split_and_save(self):
        input_file = self.input_file_entry.get()
        output_folder = self.output_folder_entry.get()
        
        if not input_file or not output_folder:
            messagebox.showerror("Error", "Please select both input file and output folder.")
            return

        try:
            # Read the Excel file
            df = pd.read_excel(input_file)
            
            # Ensure the 'Location', 'SKUCode', and 'Norm' columns are present
            if 'Location' not in df.columns or 'SKUCode' not in df.columns or 'Norm' not in df.columns:
                raise ValueError("Excel file must have 'Location', 'SKUCode', and 'Norm' columns.")

            # Rename columns
            df.rename(columns={'SKUCode': 'Partnumber', 'Norm': 'Qty'}, inplace=True)
            
            # Ensure 'Partnumber' is treated as text
            df['Partnumber'] = df['Partnumber'].astype(str)
            
            # Keep only the required columns
            df = df[['Location', 'Partnumber', 'Qty']]
            
            # Get the unique values in the 'Location' column
            unique_values = df['Location'].unique()
            
            # Create separate files for each unique value
            for value in unique_values:
                # Sanitize the value for use in filenames
                safe_value = str(value).replace('/', '_').replace('\\', '_')
                
                # Filter the data for the current unique value
                filtered_df = df[df['Location'] == value]
                
                # Define the output file path
                output_file = f"{output_folder}/{safe_value}.xlsx"
                
                # Save the filtered data to a new Excel file
                self.save_to_excel(filtered_df, output_file)
            
            messagebox.showinfo("Success", "Files have been split and saved successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def save_to_excel(self, df, file_path):
        # Create a new workbook and add a worksheet
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"  # Set the sheet name to "Sheet1"

        # Append the DataFrame header
        ws.append(df.columns.tolist())
        
        # Append the DataFrame rows
        for row in dataframe_to_rows(df, index=False, header=False):
            ws.append(row)

        # Format the 'Partnumber' column to ensure it is treated as text
        for cell in ws['B']:
            cell.number_format = '@'  # '@' format means text in Excel

        # Save the workbook to the file path
        wb.save(file_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSplitterApp(root)
    root.mainloop()
