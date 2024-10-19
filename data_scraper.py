from selenium import webdriver
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import time
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import openpyxl

# Path to your chromedriver
chromedriver_path = r'C:\\Users\\Sanjeev\\Downloads\\Temp\\chromedriver-win32\\chromedriver.exe'

# Set up the Service for ChromeDriver
service = Service(executable_path=chromedriver_path)

# Initialize the browser using the Service
driver = webdriver.Chrome(service=service)


# Open the website
print("Opening the browser...")
driver.get('https://edistrict.bih.nic.in/EduSalLink7/EDUSALARY/SchoolList.aspx')
print("Website loaded, selecting options...")


# Select district from the dropdown
select1 = Select(driver.find_element(By.ID, 'ctl00_MainContent_ddlDist'))
#select1.select_by_visible_text('BUXAR')
#select1.select_by_visible_text('ROHTAS')
select1.select_by_visible_text('KAIMUR')


# Block options
buxar_options = ['Barhampur', 'Buxar', 'Chakki', 'Chaugain', 'Chausa', 'Dumraon', 'Itarhi', 'Kesath', 'Nawanagar', 'Rajpur', 'Simri']

rohtas_options = ['Bikramganj', 'Chenari', 'Dawath', 'Dehri', 'Dinara', 'Karakat', 'Kargahar', 'Kochas', 'Nasriganj', 'Nauhatta','Nokha', 'Rajpur', 'Rohtas', 'Sanjhauli', 'Sasaram', 'Sheosagar', 'Suryapura', 'Tilouthu' ]

kaimur_options = ['Adhaura', 'Bhabua', 'Bhagwanpur', 'Chainpur', 'Chand', 'Durgawati', 'Kudra', 'Mohania', 'Nuaon', 'Ramgarh', 'Rampur']


# Set up the Excel file for appending data
excel_file = 'output.xlsx'
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Sl. No.', 'District Name', 'Block Name', 'DISE CODE', 'School Name', 'Type Of School', 'Class From', 'Class To'])


# Iterate over block options
for  idx, option in enumerate(kaimur_options):
    # Select block
    select2 = Select(driver.find_element(By.ID, 'ctl00_MainContent_ddlBlock'))
    select2.select_by_visible_text(option)

    # Submit the form
    driver.find_element(By.ID, 'ctl00_MainContent_btnSave').click()

    # Wait for the page to load (adjust time if necessary)
    time.sleep(10)

    # Parse the page with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find the table (adjust based on the actual page structure)
    table = soup.find('table')
    rows = table.find_all('tr')

    # Extract the table data
    for row in rows:
        cols = row.find_all('td')
        cols = [col.text.strip() for col in cols]

        # Append row to Excel, if the row contains data
        if cols:
           ws.append([idx + 1, 'BUXAR', option] + cols)

    # Save the workbook after each iteration to avoid data loss
    wb.save(excel_file)

# Close the browser after processing all options
driver.quit()
