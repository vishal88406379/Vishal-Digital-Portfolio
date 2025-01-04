import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import font
import os
import sys
import pandas as pd

# Function to get the path of a file
def resource_path(relative_path):
    """ Get the absolute path to the resource, works for development and for bundled .exe """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(__file__))
        return os.path.join(base_path, relative_path)
    except Exception:
        return os.path.join(os.path.dirname(__file__), relative_path)

# Paths to the Excel files (using resource_path for packaged version)
location_master_path = resource_path("All Location TATA CVBU & PCBU.xlsx")
partmaster_path = resource_path("Partmaster.xlsx")
icon_path = resource_path("scsicon.ico")

# Function to read and process the Excel files
def process_excel_files():
    try:
        # Read the Excel files using pandas
        location_data = pd.read_excel(location_master_path)
        partmaster_data = pd.read_excel(partmaster_path)

        # Example processing (you can add your logic here)
        print("Location Master Data Loaded:")
        print(location_data.head())  # Show first few rows
        print("\nPartmaster Data Loaded:")
        print(partmaster_data.head())

        messagebox.showinfo("Success", "Excel files loaded and processed successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Error processing Excel files: {e}")
        print(f"Error: {e}")

# Function to display the copyright message in the window
def display_copyright():
    copyright_text = "Copyright (c) 2024, Vishal - Spare Care Solution\n"
    copyright_text += "All rights reserved.\n"
    copyright_text += "This application is for the TATA Team."
    return copyright_text

# List of all script paths
scripts = [
    "Resrve Stock Final.py",
    "Final Pending GRN.py",
    "WIP FINAL.py"
]

# Function to run the selected script
def run_script(script):
    try:
        # Use sys._MEIPASS if running from a packaged .exe
        script_path = resource_path(script)
        print(f"\nRunning script: {script_path}")

        # Run the script and capture output
        result = subprocess.run(["python", script_path], check=True, capture_output=True, text=True)

        # If successful, show output in messagebox and status
        messagebox.showinfo("Success", f"Script '{script}' Run Successfully!")
        status_label.config(text=f"Script '{script}' Run Successfully.", fg="green")
        print(f"Script Output: {result.stdout}")  # Print output in terminal

    except subprocess.CalledProcessError as e:
        # Handle errors by displaying detailed output
        messagebox.showerror("Error", f"Error running script: {e}\nExit code: {e.returncode}\n{e.stderr}")
        status_label.config(text=f"Error: {e.stderr}", fg="red")
        print(f"Error: {e.stderr}")  # Print error message

# Function to handle the script selection from the dropdown
def on_menu_select(event=None):
    selected_script = script_menu.get()
    if selected_script:
        if selected_script in scripts:
            run_script(selected_script)

# Creating the main window
root = tk.Tk()
root.title("TATA Team Script Runner")
root.config(bg="#f0f0f0")

# Set the window icon
root.iconbitmap(icon_path)  # Path to your icon file

# Displaying the copyright message in the window
copyright_text = display_copyright()

# Adding a header label
header_font = font.Font(family="Helvetica", size=16, weight="bold")
header_label = tk.Label(root, text="Spare Care Solution", font=header_font, bg="#f0f0f0", fg="#333")
header_label.pack(pady=10)

# Adding a label to display your name
name_label = tk.Label(root, text="Vishal", font=("Helvetica", 14), bg="#f0f0f0", fg="#333")
name_label.pack(pady=5)

# Displaying copyright message
copyright_label = tk.Label(root, text=copyright_text, font=("Helvetica", 10), bg="#f0f0f0", fg="#555")
copyright_label.pack(pady=5)

# Displaying your phone number
phone_label = tk.Label(root, text="Contact: 9129572268", font=("Helvetica", 10), bg="#f0f0f0", fg="#555")
phone_label.pack(pady=5)

# Adding a label to prompt for input
label = tk.Label(root, text="Select a script from the menu below:", bg="#f0f0f0")
label.pack(pady=10)

# Dropdown menu to select the script (showing actual script file names)
script_menu = ttk.Combobox(root, values=scripts, state="readonly")
script_menu.pack(pady=10, padx=20, fill='x')

# Add a button to run the selected script
run_button = tk.Button(root, text="Run Selected Script", width=20, height=2, bg="#4CAF50", fg="white", command=on_menu_select)
run_button.pack(pady=20)

# Add a "Exit" button
exit_button = tk.Button(root, text="Exit", width=20, bg="#f44336", fg="white", command=root.quit)
exit_button.pack(pady=10)

# Status label to display script execution status
status_label = tk.Label(root, text="Select a script and press Run", bg="#f0f0f0", fg="black")
status_label.pack(pady=10)

# Add a button to process the Excel files
excel_button = tk.Button(root, text="Process Excel Files", width=20, height=2, bg="#2196F3", fg="white", command=process_excel_files)
excel_button.pack(pady=20)

# Configure the window to be resizable and auto-fit based on content
root.resizable(True, True)

# Run the Tkinter event loop
root.mainloop()
