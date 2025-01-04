import os

def create_folders():
    # List of dealer names
    dealers = [
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

    # List of subfolders to create under each dealer
    subfolders = [
        "OTC INVOICE",
        "PURCHASE LINE PO",
        "SPARE CONSUMPTION",
        "CLOSING STOCK",
        "stock transaction",
        "Job Line Invoice"
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
