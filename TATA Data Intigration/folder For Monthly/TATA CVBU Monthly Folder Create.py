import os

def create_folders():
    # List of dealer names
    dealers = [
        "Anand Trucking",
        "ANUPAM MOTORS",
        "Asthavinayak Motors",
        "AUTO XPERTS",
        "Bhandari Auto",
        "Binod Auto CVBU",
        "Budhia Agencies",
        "CK Motors",
        "CROSSLAND TRUCKS",
        "DADA Motors CVBU",
        "DIAMOND WHEELS",
        "DURGA AUTOMART",
        "ENAR INDUSTRIAL",
        "ESTEE DEALERS",
        "EXCEL VEHICLES",
        "Gajraj Vahan",
        "Ganganagar Motors",
        "GNG Auto",
        "Ideal Dealers",
        "INFINITY MOTORS",
        "Johar Automobiles",
        "KALINGA AUTO",
        "KKRISHNA VAAHAN",
        "KOHLI & SONS",
        "Krishna Autotech",
        "LAKSH AGENTS",
        "Lexus Motors",
        "Libra Automobiles",
        "LOTUS MOTORS",
        "Matsya Comm",
        "MG Motors",
        "OM AUTOWHEELS",
        "OSL Automotives",
        "Paraj Motors",
        "Raj Enterprises",
        "RATHI AUTO",
        "RH Automobiles",
        "Roshan CVBU",
        "SAI SUDHA MOTORS",
        "Samal Auto",
        "SKS MOTORS",
        "SM AUTO SALES",
        "SM Auto",
        "SREELN MOTORS",
        "Triumph 2080480",
        "Triumph 2080482 TASS",
        "Triumph Auto CV",
        "Triumph SCV 2087760",
        "Triumph TASS Workshop",
        "Trupti Automotives",
        "VASUNDHARA MOTORS",
        "VIKRAMSHILA AUTO",
        "VINAYAK AUTO"
    ]

    # List of subfolders to create under each dealer
    subfolders = [
        "OTC INVOICE",
        "PURCHASE LINE PO",
        "Purchase Line Items",
        "SPARE CONSUMPTION",
        "CLOSING STOCK",
        "stock transaction",
        "channel partner",
        "Job Line Invoice",
        "sap purchase order reason"
    ]

    # Create dealer folders and subfolders
    for dealer in dealers:
        # Create main dealer folder
        os.makedirs(dealer, exist_ok=True)
        
        # Create subfolders within the dealer folder
        for subfolder in subfolders:
            os.makedirs(os.path.join(dealer, subfolder), exist_ok=True)

    print("Folders created successfully.")

if __name__ == "__main__":
    create_folders()
