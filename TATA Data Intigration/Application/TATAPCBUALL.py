# SCS Vishal
# Contact: 9129572268
# Copyright © 2024 Vishal. All Rights Reserved.

import subprocess

# List of your script paths and their names for the menu
scripts = [
    ("1. Compilation Script", "C:/Users/Vishal/Desktop/TATA PCBU/compilation.py"),
    ("2. OeminvoiceTATACVBU", "C:/Users/Vishal/Desktop/TATA PCBU/OeminvoiceTATACVBU.py"),
    ("3. OeminvoiceTATAPCBU", "C:/Users/Vishal/Desktop/TATA PCBU/OeminvoiceTATAPCBU.py"),
    ("4. OrdersentforONEspecific", "C:/Users/Vishal/Desktop/TATA PCBU/OrdersentforONEspecific.py"),
    ("5. PolocationwiseTATA&CVBU", "C:/Users/Vishal/Desktop/TATA PCBU/PolocationwiseTATA&CVBU.py"),
    ("6. Stock&ReservestockPendingGRN", "C:/Users/Vishal/Desktop/TATA PCBU/Stock&ReservestockPendingGRN.py"),
    ("7. Stockincreaseformating", "C:/Users/Vishal/Desktop/TATA PCBU/Stockincreaseformating.py"),
    ("8. StockuploadTATACV&PV", "C:/Users/Vishal/Desktop/TATA PCBU/StockuploadTATACV&PV.py")
]

def run_script(script_path):
    """Function to run the selected script."""
    subprocess.run(["python", script_path])

def show_menu():
    """Function to show the main menu."""
    print("\nSCS Vishal")
    print("Contact: 9129572268")
    print("Copyright © 2024 Vishal. All Rights Reserved.\n")
    
    print("Please select the script you want to run:")
    for script in scripts:
        print(script[0])
    print("0. Exit")

def main():
    """Main function to handle menu and script selection."""
    while True:
        show_menu()
        
        # Ask the user for their choice
        try:
            choice = int(input("\nEnter the number of the script you want to run (1-8), or 0 to Exit: "))
        except ValueError:
            print("Invalid input, please enter a number between 0 and 8.")
            continue
        
        # If the user wants to exit, break the loop
        if choice == 0:
            print("Exiting the application.")
            break
        
        # Ensure the choice is valid
        if 1 <= choice <= 8:
            selected_script = scripts[choice - 1][1]
            print(f"\nRunning {scripts[choice - 1][0]}...\n")
            run_script(selected_script)
        else:
            print("Invalid choice. Please enter a number between 1 and 8.")

# Start the program
if __name__ == "__main__":
    main()
