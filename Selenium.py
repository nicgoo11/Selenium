from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

options = Options()

# setup the webdriver
driver_path = "C:\\tools\\chromedriver.exe"

service = Service(driver_path)
driver = webdriver.Chrome(service=service, options=options)


driver.get('https://www.egolzwil.ch/anlaesseaktuelles?ort=')

# to store the scraped data
data = []

# iterate over each page
for page in range(1, 5):
    # Wait
    wait = WebDriverWait(driver, 20)
    table = wait.until(EC.presence_of_element_located((By.ID, 'anlassList')))

    # locate the table and its rows
    rows = table.find_elements(By.TAG_NAME, 'tr')

    # iterate over each row
    for row in rows:
        try:
            date = row.find_element(By.CLASS_NAME, 'icms-datatable-col-_anlassDate').text
        except NoSuchElementException:
            date = None

        # extract the topic
        try:
            topic = row.find_element(By.TAG_NAME, 'a').text
        except NoSuchElementException:
            topic = None

        # extract the location
        try:
            location_element = row.find_element(By.CSS_SELECTOR, '.icms-datatable-col-lokalitaet')
            location = driver.execute_script("return arguments[0].textContent;", location_element)
        except NoSuchElementException:
            location = None

        try:
            organiser_element = row.find_element(By.CLASS_NAME, 'icms-datatable-col-organisator')
            organiser = driver.execute_script("return arguments[0].innerText;", organiser_element)
        except NoSuchElementException:
            organiser = None

        # store the data
        data.append({
            'date': date,
            'topic': topic,
            'location': location,
            'organiser': organiser,
        })

    # navigate to the next page
    next_page_link = driver.find_element(By.XPATH, f'//a[@data-dt-idx="{page + 1}"]')
    next_page_link.click()

# close the driver
driver.close()

# print the data
for item in data:
    print(item)

# write the data to excel
import pandas as pd

df = pd.DataFrame(data)

# create a new column combining 'topic' and 'organiser'
df['topic - organiser'] = df['topic'] + ' - ' + df['organiser']

# remove rows where 'topic' contains certain strings
df = df[~df['topic'].str.contains('grünabfuhr', case=False, na=False)]
df = df[~df['topic'].str.contains('ferien', case=False, na=False)]
df = df[~df['topic'].str.contains('ütter- und Väterberatu', case=False, na=False)]

df = df[['topic - organiser', 'location', 'date']]

# write to excel
df.to_excel("data.xlsx", index=False)
