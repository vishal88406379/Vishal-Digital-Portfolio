import os
import shutil

# Define the folder path where files are located
folder_path = r"C:\Users\Vishal\OneDrive\Desktop\sisent december"

# Function to extract dealer name from the file name (third part after splitting by '_')
def extract_dealer_name(file_name):
    parts = file_name.split('_')
    if len(parts) >= 3:
        # The third part is the dealer name
        return parts[2]
    return "UnknownDealer"  # Fallback if no dealer name found

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

# Create folders for each unique dealer name
for dealer in dealer_set:
    dealer_folder = os.path.join(folder_path, dealer)
    if not os.path.exists(dealer_folder):
        os.makedirs(dealer_folder)
        print(f"Created folder for dealer: {dealer}")
    else:
        print(f"Folder already exists for dealer: {dealer}")

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

print("Files organized successfully.")
