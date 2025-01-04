import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font

def log_error(message):
    """Logs errors to an error log file."""
    with open("error_log.txt", "a") as f:
        f.write(f"{pd.Timestamp.now()}: {message}\n")

def compile_excel_files(folder_path):
    """Compiles all Excel files from a folder into a DataFrame."""
    try:
        files = os.listdir(folder_path)
    except FileNotFoundError:
        log_error(f"The folder path '{folder_path}' does not exist.")
        return pd.DataFrame(), f"Folder path does not exist: {folder_path}"
    except PermissionError:
        log_error(f"Permission denied for folder path '{folder_path}'.")
        return pd.DataFrame(), f"Permission denied: {folder_path}"

    excel_files = [f for f in files if f.endswith('.xlsx')]
    compiled_df = pd.DataFrame()

    if not excel_files:
        return compiled_df, "No Excel files found in the folder."

    for file in excel_files:
        input_file = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(input_file, sheet_name=None)  # Read all sheets
            for sheet_name, sheet_data in df.items():
                compiled_df = pd.concat([compiled_df, sheet_data], ignore_index=True)
        except Exception as e:
            log_error(f"Error processing file {file}: {e}")

    return compiled_df, None

def run_compile():
    """Runs the compilation process for all dealer folders."""
    main_folder = folder_entry.get()
    if not main_folder:
        messagebox.showwarning("Warning", "Please select the main dealer folder.")
        return

    results = []

    dealer_names = [
        "OTC INVOICE", "PURCHASE LINE PO", "Purchase Line Items",
        "SPARE CONSUMPTION", "CLOSING STOCK", "stock transaction",
        "channel partner", "Job Line Invoice", "sap purchase order reason"
    ]

    for dealer in dealer_names:
        dealer_folder = os.path.join(main_folder, dealer)
        if os.path.isdir(dealer_folder):
            compiled_df, error = compile_excel_files(dealer_folder)
            if compiled_df.empty:
                results.append(f"{dealer}: {error if error else 'No data to compile.'}")
                continue
            
            output_file = os.path.join(main_folder, f"{dealer}.xlsx")
            try:
                wb = Workbook()
                ws = wb.active

                for r in dataframe_to_rows(compiled_df, index=False, header=True):
                    ws.append(r)

                # Adjust column widths and make header bold
                for column in ws.columns:
                    max_length = max(len(str(cell.value)) for cell in column if cell.value)
                    adjusted_width = (max_length + 2)
                    ws.column_dimensions[column[0].column_letter].width = adjusted_width

                for cell in ws[1]:
                    cell.font = Font(bold=True)

                wb.save(output_file)
                results.append(f"{dealer}: File saved successfully as {output_file}.")
            except Exception as e:
                log_error(f"Error saving file for {dealer}: {e}")
                results.append(f"{dealer}: Error saving file: {e}")
        else:
            results.append(f"{dealer}: Skipped (folder not found).")

    final_message = "\n".join(results)
    messagebox.showinfo("Processing Complete", f"Results:\n\n{final_message}")

# Create the main window
root = tk.Tk()
root.title("TATA SCS Excel File Compiler")

frame_header = tk.Frame(root, padx=10, pady=10)
frame_header.pack(fill=tk.X)

frame_buttons = tk.Frame(root, padx=10, pady=10)
frame_buttons.pack(fill=tk.X)

header_label = tk.Label(frame_header, text="TATA SCS Monthly Data Compiler", font=("Arial", 16, "bold"))
header_label.pack()

folder_entry = tk.Entry(frame_buttons, width=40)
folder_entry.pack(padx=5, pady=5)

browse_folder_button = tk.Button(frame_buttons, text="Browse Main Dealer Folder", command=lambda: folder_entry.insert(0, filedialog.askdirectory()), bg="lightgreen")
browse_folder_button.pack()

run_all_button = tk.Button(frame_buttons, text="Run All", command=run_compile, bg="blue", fg="white", font=('Arial', 12, 'bold'))
run_all_button.pack()

# Start the GUI event loop
root.mainloop()
