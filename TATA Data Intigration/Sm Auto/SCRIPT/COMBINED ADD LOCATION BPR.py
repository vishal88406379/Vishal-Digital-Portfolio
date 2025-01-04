import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class FileCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Processor with Location Mapping")

        # Input files
        self.files_label = tk.Label(root, text="Select Excel Files:")
        self.files_label.pack(pady=5)
        
        self.files_entry = tk.Entry(root, width=50)
        self.files_entry.pack(pady=5)
        
        self.browse_files_button = tk.Button(root, text="Browse", command=self.browse_files)
        self.browse_files_button.pack(pady=5)

        # Location mapping file
        self.mapping_label = tk.Label(root, text="Select Location Mapping File:")
        self.mapping_label.pack(pady=5)
        
        self.mapping_entry = tk.Entry(root, width=50)
        self.mapping_entry.pack(pady=5)
        
        self.browse_mapping_button = tk.Button(root, text="Browse", command=self.browse_mapping_file)
        self.browse_mapping_button.pack(pady=5)

        # Output file
        self.output_label = tk.Label(root, text="Select Output File:")
        self.output_label.pack(pady=5)
        
        self.output_entry = tk.Entry(root, width=50)
        self.output_entry.pack(pady=5)
        
        self.browse_output_button = tk.Button(root, text="Browse", command=self.browse_output_file)
        self.browse_output_button.pack(pady=5)
        
        # Run button
        self.run_button = tk.Button(root, text="Process Files", command=self.process_files)
        self.run_button.pack(pady=20)
        
    def browse_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if file_paths:
            self.files_entry.delete(0, tk.END)
            self.files_entry.insert(0, ';'.join(file_paths))

    def browse_mapping_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.mapping_entry.delete(0, tk.END)
            self.mapping_entry.insert(0, file_path)

    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, file_path)

    def process_files(self):
        file_paths = self.files_entry.get().split(';')
        mapping_file = self.mapping_entry.get()
        output_file = self.output_entry.get()
        
        if not file_paths or not mapping_file or not output_file:
            messagebox.showerror("Error", "Please select input files, location mapping file, and output file.")
            return
        
        try:
            # Read the location mapping file
            mapping_df = pd.read_excel(mapping_file)
            
            # Ensure 'Sold_To_Party' and 'Location' columns are present
            if 'Sold_To_Party' not in mapping_df.columns or 'Location' not in mapping_df.columns:
                raise ValueError("Location mapping file must have 'Sold_To_Party' and 'Location' columns.")
            
            # Initialize an empty list to hold DataFrames
            dataframes = []
            
            # Process each file
            for file_path in file_paths:
                df = pd.read_excel(file_path)
                
                # Ensure 'Sold_To_Party' column is present
                if 'Sold_To_Party' not in df.columns:
                    raise ValueError(f"File {file_path} does not have 'Sold_To_Party' column.")
                
                # Merge with mapping to add 'Location'
                df = df.merge(mapping_df, on='Sold_To_Party', how='left')
                
                # Append to the list of DataFrames
                dataframes.append(df)
            
            # Combine all DataFrames into a single DataFrame
            combined_df = pd.concat(dataframes, ignore_index=True)
            
            # Save the combined data to a new Excel file
            combined_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", "Files have been processed and saved successfully.")
        
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = FileCombinerApp(root)
    root.mainloop()
