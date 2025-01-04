import pandas as pd
import glob
from tkinter import Tk, Label, Button, StringVar, filedialog, messagebox
from openpyxl import load_workbook
import os

def process_files():
    try:
        # Get the input folder and output folder paths
        input_folder = folder_path.get()
        output_folder = file_path.get()
        location_file = location_file_path.get()

        if not input_folder:
            messagebox.showerror("Error", "No folder selected.")
            return
        
        if not output_folder:
            messagebox.showerror("Error", "No output folder selected.")
            return
        
        if not location_file:
            messagebox.showerror("Error", "No location file selected.")
            return

        # Load location data
        try:
            location_df = pd.read_excel(location_file)
            if 'Sold_To_Party' not in location_df.columns or 'Code' not in location_df.columns or 'Location' not in location_df.columns:
                messagebox.showerror("Error", "Location file must contain 'Sold_To_Party', 'Code', and 'Location' columns.")
                return
        except Exception as e:
            messagebox.showerror("Error", f"Error loading location file: {e}")
            return

        # Locate all Excel files in the selected folder
        files = glob.glob(f'{input_folder}/*.xlsx')

        if not files:
            messagebox.showerror("Error", "No Excel files found in the selected folder.")
            return

        # Initialize an empty DataFrame to hold all data
        all_data = pd.DataFrame()

        # Loop through each file and process it
        for file in files:
            try:
                df = pd.read_excel(file)
                # Check if required columns exist in the DataFrame
                if 'SKUCode' in df.columns and 'Norm' in df.columns and 'Sold_To_Party' in df.columns:
                    # Select only the required columns and rename them
                    filtered_df = df[['SKUCode', 'Norm', 'Sold_To_Party']].copy()
                    filtered_df.columns = ['Partnumber', 'QTY', 'Sold_To_Party']
                    
                    # Convert 'Partnumber' column to text using .loc
                    filtered_df.loc[:, 'Partnumber'] = filtered_df['Partnumber'].astype(str)
                    
                    # Append the filtered DataFrame to the main DataFrame
                    filtered_df['Original_File'] = os.path.basename(file)  # Add a column with the original file name
                    all_data = pd.concat([all_data, filtered_df], ignore_index=True)
                else:
                    print(f"Required columns not found in {file}. Skipping this file.")
            except Exception as e:
                print(f"Error processing file {file}: {e}")

        if 'Sold_To_Party' not in all_data.columns:
            messagebox.showerror("Error", "Column 'Sold_To_Party' not found in the data.")
            return

        # Merge with location data
        all_data = pd.merge(all_data, location_df, on='Sold_To_Party', how='left')

        if all_data.empty:
            messagebox.showerror("Error", "No matching data found after merging with location data.")
            return

        # Process and save data based on location
        location_values = all_data['Location'].dropna().unique()
        original_files = all_data['Original_File'].unique()

        for location in location_values:
            # Filter data where 'Location' matches the current value
            matched_data = all_data[all_data['Location'] == location]

            if matched_data.empty:
                continue
            
            # Define the output file path for each location
            match_output_file = f'{output_folder}/{location}_matched_data.xlsx'
            
            # Select only 'Partnumber' and 'QTY' columns
            output_data = matched_data[['Partnumber', 'QTY']]
            
            output_data.to_excel(match_output_file, index=False)
            adjust_column_width(match_output_file)

        # Handle unmatched data
        unmatched_data = all_data[all_data['Location'].isna()]

        for original_file in original_files:
            unmatched_file_data = unmatched_data[unmatched_data['Original_File'] == original_file]

            if unmatched_file_data.empty:
                continue
            
            # Define the output file path for unmatched data
            original_file_name = os.path.splitext(original_file)[0]  # Remove extension from original file name
            unmatched_output_file = f'{output_folder}/{original_file_name}_unmatched_data.xlsx'
            
            # Select only 'Partnumber' and 'QTY' columns
            output_data = unmatched_file_data[['Partnumber', 'QTY']]
            
            output_data.to_excel(unmatched_output_file, index=False)
            adjust_column_width(unmatched_output_file)

        messagebox.showinfo("Success", "Data extraction complete. Files saved to the selected output folder.")

    except Exception as e:
        messagebox.showerror("Error", f"An unexpected error occurred: {e}")

def adjust_column_width(file_path):
    try:
        # Load the workbook and select the active sheet
        workbook = load_workbook(file_path)
        sheet = workbook.active

        # Iterate through all columns and set the width based on the maximum length of the content
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = max_length + 2  # Add a little extra space
            sheet.column_dimensions[column_letter].width = adjusted_width

        # Save the adjusted workbook
        workbook.save(file_path)
    except Exception as e:
        print(f"Error adjusting column width for {file_path}: {e}")

# Function to select input folder
def select_folder():
    folder_selected = filedialog.askdirectory(title='Select Folder Containing Excel Files')
    folder_path.set(folder_selected)

# Function to select output folder
def select_output_folder():
    folder_selected = filedialog.askdirectory(title='Select Output Folder')
    file_path.set(folder_selected)

# Function to select location file
def select_location_file():
    file_selected = filedialog.askopenfilename(title='Select Location File', filetypes=[('Excel Files', '*.xlsx')])
    location_file_path.set(file_selected)

# Initialize Tkinter
root = Tk()
root.title("Excel Data Processor")

# Variables to hold folder paths
folder_path = StringVar()
file_path = StringVar()
location_file_path = StringVar()

# Create GUI elements
Label(root, text="Select Folder Containing Excel Files:").pack(pady=5)
Button(root, text="Browse", command=select_folder).pack(pady=5)
Label(root, textvariable=folder_path).pack(pady=5)

Label(root, text="Select Output Folder:").pack(pady=5)
Button(root, text="Browse", command=select_output_folder).pack(pady=5)
Label(root, textvariable=file_path).pack(pady=5)

Label(root, text="Select Location File:").pack(pady=5)
Button(root, text="Browse", command=select_location_file).pack(pady=5)
Label(root, textvariable=location_file_path).pack(pady=5)

Button(root, text="Run", command=process_files).pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
