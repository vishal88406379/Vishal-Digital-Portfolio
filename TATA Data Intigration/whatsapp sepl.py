from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import pyperclip
import time

# Configure webdriver
driver = webdriver.Chrome()
driver.maximize_window()
driver.get("https://web.whatsapp.com/")
time.sleep(30)
wait = WebDriverWait(driver, 120)

# Read the Excel file
df = pd.read_excel(r"C:\Users\HP\Desktop\all data\whatstest\vishalmsg.xlsx")

# Check if the DataFrame is empty
if df.empty:
    print("No data available in the Excel file.")
else:
    # Iterate over each row in the DataFrame
    for _, row in df.iterrows():
        name = row['Name']
        phone_number = str(int(row['Phone Number']))  # Ensure the phone number is a string

        # Job description message
        message = f"""
Hello {name},

**Job Opportunity Alert**

**Job Title**: Data Entry Operator  
**Company**: Superb Enterprises Pvt Ltd  
**Location**: ITO 2nd Floor Nehru House, 110002 (Near ITO Metro Gate No-4)  
**Experience**: 1-3 years  

**Job Summary**:  
Seeking an Operation Executive with 1-3 years of experience in managing operations, data entry, or similar roles. Tasks include preparing reports, working on MS Excel, Google Sheets, and more data management software.  

**Responsibilities**:  
- Prepare and maintain daily reports  
- Manage emails and communicate effectively  
- Documentation & Excel  
- Typing speed  

**Qualifications**:  
- Graduation  
- 1-3 years in operations  
- Strong organizational and communication skills  
- Proficiency in Microsoft Office  

**Company Websites**:  
- [Superb Attestation](https://superbattestation.com/about-us.asp)  
- [Superb My Trip](https://www.superbmytrip.com/)  
- [Superb Study Abroad](https://superbstudyabroad.com/about-us)  
- [Umrah Services](https://www.umrahservices.in/)  

Best regards,  
Superb Enterprises Pvt Ltd
"""

        # Copy the message to clipboard
        pyperclip.copy(message)

        # Locate the search box in WhatsApp Web
        search_xpath = '//div[@contenteditable="true"][@data-tab="3"]'
        search_box = WebDriverWait(driver, 120).until(EC.visibility_of_element_located((By.XPATH, search_xpath)))

        try:
            search_box.click()
            time.sleep(2)
            search_box.send_keys(Keys.CONTROL + "a")  # Select all text
            search_box.send_keys(Keys.BACKSPACE)  # Clear the search box
        except Exception as e:
            print(f"Error clearing search box: {e}")
            continue

        # Search for the contact by phone number
        search_box.send_keys(phone_number + Keys.ENTER)
        time.sleep(2)

        # Verify if the contact is found
        try:
            WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, "//span[@class='_ao3e']")))
        except:
            print(f"Contact not found for this phone number: {phone_number}")
            search_box.clear()
            continue

        # Send the message
        message_box_xpath = '//div[@aria-placeholder="Type a message"]'
        message_box = wait.until(EC.visibility_of_element_located((By.XPATH, message_box_xpath)))
        message_box.click()

        # Paste the copied text
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        ActionChains(driver).send_keys(Keys.RETURN).perform()

        # Clear the search box for the next contact
        search_box.clear()
        time.sleep(2)

print("Messages sent to all contacts. Close the browser manually if needed.")
