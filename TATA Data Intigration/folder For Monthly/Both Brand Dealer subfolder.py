import os

def create_folders():
    # List of dealer names for TATA PCBU
    tata_pcb_dealers = [
        "Adishakti cars pvt ltd",
        "AG MOTORS",
        "AKAR FOURWHEEL",
        "ANAND MOTOREN PRIVATE LIMITED",
        "Anjney Auto pvt Ltd",
        "Autoprime",
        "Bhasin Motors",
        "Bimal Cars",
        "Basudeb Auto Ltd",
        "Classic Motors",
        "Dada Motors PCBU",
        "Ganganagar Automobiles Pvt Ltd",
        "Go Auto",
        "Himgiri Automobiles",
        "Ideal Dealers Private Limited",
        "JD Autonation",
        "Keshva Motors",
        "Krishna Car World",
        "Lexican Motors",
        "Lexus Motors",
        "Marudhar Motors",
        "Multitech Motors",
        "National Garage",
        "Planet Spares",
        "Ravindra Auto",
        "Rising Auto",
        "Roshan PCBU",
        "Seth and Sons",
        "Shree ji Automart pvt Ltd",
        "SHREE SHYAM MOTORS",
        "SHRI VASU AUTOMOBILES LTD",
        "Smam Automart",
        "STELLAR AUTODRIVE",
        "Triumph PCBU",
        "TRUENORTH AUTOMOBILES",
        "Zedex Motors",
        "Binod Auto PCBU",
        "KD Motor"
    ]

    # List of dealer names for TATA CVBU
    tata_cv_dealers = [
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

    # Create TATA PCBU folder and its dealer folders
    os.makedirs("TATA PCBU", exist_ok=True)
    for dealer in tata_pcb_dealers:
        dealer_path = os.path.join("TATA PCBU", dealer)
        os.makedirs(dealer_path, exist_ok=True)
        for subfolder in subfolders:
            os.makedirs(os.path.join(dealer_path, subfolder), exist_ok=True)

    # Create TATA CVBU folder and its dealer folders
    os.makedirs("TATA CVBU", exist_ok=True)
    for dealer in tata_cv_dealers:
        dealer_path = os.path.join("TATA CVBU", dealer)
        os.makedirs(dealer_path, exist_ok=True)
        for subfolder in subfolders:
            os.makedirs(os.path.join(dealer_path, subfolder), exist_ok=True)

    print("Folders created successfully.")

if __name__ == "__main__":
    create_folders()
