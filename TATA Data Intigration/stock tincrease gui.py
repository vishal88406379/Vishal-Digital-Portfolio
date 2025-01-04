import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

# Function to extract dealer name from the file name (third part after splitting by '_')
def extract_dealer_name(file_name):
    parts = file_name.split('_')
    if len(parts) >= 3:
        # The third part is the dealer name
        return parts[2]
    return "UnknownDealer"  # Fallback if no dealer name found

# Function to organize files
def organize_files():
    folder_path = folder_path_entry.get()
    
    if not folder_path:
        messagebox.showerror("Error", "Please select a folder.")
        return

    if not os.path.exists(folder_path):
        messagebox.showerror("Error", "The selected folder does not exist.")
        return
    
    # Set to store unique dealer names
    dealer_set = set()

    # First pass: Identify dealer names from file names
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        # Skip directories
        if os.path.isdir(file_path):
            continue

        # Extract dealer name
        dealer_name = extract_dealer_name(file_name)
        dealer_set.add(dealer_name)

    # Create folders for each unique dealer name with number series
    for dealer in dealer_set:
        dealer_folder = os.path.join(folder_path, dealer)
        
        if os.path.exists(dealer_folder):
            # If the folder already exists, create a new folder with a number suffix
            counter = 1
            while os.path.exists(dealer_folder + f"_{counter}"):
                counter += 1
            dealer_folder = dealer_folder + f"_{counter}"

        os.makedirs(dealer_folder)
        print(f"Created folder for dealer: {dealer_folder}")
    
    # Second pass: Copy files into their respective dealer folders
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)

        # Skip directories
        if os.path.isdir(file_path):
            continue

        # Extract dealer name and copy file
        dealer_name = extract_dealer_name(file_name)
        destination_path = os.path.join(folder_path, dealer_name, file_name)

        # Copy file to the dealer's folder
        shutil.copy(file_path, destination_path)
        print(f"Copied: {file_name} to {dealer_name} folder")

    messagebox.showinfo("Success", "Files organized successfully.")

# Create the main window
root = tk.Tk()
root.title("File Organizer")

# Create the folder path input and browse button
folder_path_label = tk.Label(root, text="Select Folder:")
folder_path_label.pack(pady=5)

folder_path_entry = tk.Entry(root, width=50)
folder_path_entry.pack(pady=5)

def browse_folder():
    folder_selected = filedialog.askdirectory()
    folder_path_entry.delete(0, tk.END)
    folder_path_entry.insert(0, folder_selected)

browse_button = tk.Button(root, text="Browse", command=browse_folder)
browse_button.pack(pady=5)

# Create the Run button
run_button = tk.Button(root, text="Run", command=organize_files)
run_button.pack(pady=20)

# Run the Tkinter event loop
root.mainloop()
