import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime

class LocationMapperApp:
    def __init__(self, master):
        self.master = master
        master.title("Add Location Mapping")
        master.geometry("400x300")

        # Base folder selection
        self.base_folder_path = ""
        self.base_folder_label = tk.Label(master, text="Select Base Folder:")
        self.base_folder_label.pack(pady=10)
        self.base_folder_button = tk.Button(master, text="Browse", command=self.select_base_folder)
        self.base_folder_button.pack(pady=5)

        # Location mapping file selection
        self.location_file_path = ""
        self.location_file_label = tk.Label(master, text="Select Location Mapping File:")
        self.location_file_label.pack(pady=10)
        self.location_file_button = tk.Button(master, text="Browse", command=self.select_location_file)
        self.location_file_button.pack(pady=5)

        # Output folder selection
        self.output_folder_path = ""
        self.output_folder_label = tk.Label(master, text="Select Output Folder:")
        self.output_folder_label.pack(pady=10)
        self.output_folder_button = tk.Button(master, text="Browse", command=self.select_output_folder)
        self.output_folder_button.pack(pady=5)

        # Run button
        self.run_button = tk.Button(master, text="Run", command=self.combine_files)
        self.run_button.pack(pady=20)

    def select_base_folder(self):
        self.base_folder_path = filedialog.askdirectory(title="Select the Base Folder")
        self.base_folder_label.config(text=f"Base Folder: {self.base_folder_path.split('/')[-1]}")

    def select_location_file(self):
        self.location_file_path = filedialog.askopenfilename(title="Select the Location Mapping File", filetypes=[("Excel files", "*.xlsx")])
        self.location_file_label.config(text=f"Location Mapping File: {self.location_file_path.split('/')[-1]}")

    def select_output_folder(self):
        self.output_folder_path = filedialog.askdirectory(title="Select Output Folder")
        self.output_folder_label.config(text=f"Output Folder: {self.output_folder_path.split('/')[-1]}")

    def generate_output_file_name(self, base_name):
        count = 1
        output_file_name = f"{base_name}.xlsx"
        output_file_path = os.path.join(self.output_folder_path, output_file_name)

        while os.path.exists(output_file_path):
            output_file_name = f"{base_name} {count}.xlsx"
            output_file_path = os.path.join(self.output_folder_path, output_file_name)
            count += 1

        return output_file_path

    def combine_files(self):
        try:
            if not self.base_folder_path or not self.location_file_path or not self.output_folder_path:
                messagebox.showwarning("Warning", "Please select all files and the output folder.")
                return

            base_files = [f for f in os.listdir(self.base_folder_path) if f.endswith('.xlsx') and not f.startswith('~')]
            if not base_files:
                messagebox.showwarning("Warning", "No Excel files found in the selected folder.")
                return

            combined_data = pd.DataFrame()
            location_data = pd.read_excel(self.location_file_path)

            for base_file in base_files:
                base_data = pd.read_excel(os.path.join(self.base_folder_path, base_file))

                # Check for duplicates in base data before merging
                if base_data.duplicated().any():
                    print(f"Duplicate rows found in {base_file}:")
                    print(base_data[base_data.duplicated()])

                # Merge data
                merged_data = base_data.merge(location_data[['Code', 'Final Location']], left_on='Division', right_on='Code', how='left')
                merged_data.rename(columns={'Final Location': 'Location'}, inplace=True)
                merged_data = merged_data.drop(columns=['Code'], errors='ignore')

                # Remove duplicates after merging
                merged_data = merged_data.drop_duplicates()

                # Ensure 'Part No' is treated as string
                if 'Part No' in merged_data.columns:
                    merged_data['Part No'] = merged_data['Part No'].astype(str)

                combined_data = pd.concat([combined_data, merged_data], ignore_index=True)

            output_file_path = self.generate_output_file_name("Sm Auto BackOrder All Location")

            with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                combined_data.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']

                # Set 'Part No' column to text format
                if 'Part No' in combined_data.columns:
                    part_no_col_idx = combined_data.columns.get_loc('Part No')
                    text_format = writer.book.add_format({'num_format': '@'})  # Text format
                    worksheet.set_column(part_no_col_idx, part_no_col_idx, None, text_format)

                # Auto-fit columns based on content width
                for idx, col in enumerate(combined_data.columns):
                    max_length = max(combined_data[col].astype(str).apply(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_length)

            # Now create the filtered file
            self.create_filtered_file(combined_data)

            messagebox.showinfo("Success", f"Combined file created successfully as '{os.path.basename(output_file_path)}'!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

    def create_filtered_file(self, combined_data):
        try:
            # Filtered data based on criteria
            # Keep all orders but only SAP-000 orders with Days Pending <= 3
            filtered_data = combined_data[
                (combined_data['Order Number'].str.startswith("SAP-000", na=False) & (combined_data['Days Pending'] <= 3)) |
                (~combined_data['Order Number'].str.startswith("SAP-000", na=False))
            ]

            # Remove records where 'Order Date' is before or on 13 Oct 2024
            filtered_data = filtered_data[pd.to_datetime(filtered_data['Order Date'], errors='coerce') > '2024-10-13']

            filtered_file_path = self.generate_output_file_name("Filtered BackOrder Data")

            with pd.ExcelWriter(filtered_file_path, engine='xlsxwriter') as writer:
                filtered_data.to_excel(writer, index=False, sheet_name='Sheet1')
                worksheet = writer.sheets['Sheet1']

                # Set 'Part No' column to text format
                if 'Part No' in filtered_data.columns:
                    part_no_col_idx = filtered_data.columns.get_loc('Part No')
                    text_format = writer.book.add_format({'num_format': '@'})  # Text format
                    worksheet.set_column(part_no_col_idx, part_no_col_idx, None, text_format)

                # Auto-fit columns based on content width
                for idx, col in enumerate(filtered_data.columns):
                    max_length = max(filtered_data[col].astype(str).apply(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_length)

            messagebox.showinfo("Success", f"Filtered file created successfully as '{os.path.basename(filtered_file_path)}'!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating the filtered file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LocationMapperApp(root)
    root.mainloop()
