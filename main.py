from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time, openpyxl

# Load the Excel file
excel_file_path = "/Users/md.mahmudurrahman/PycharmProjects/Tec Pack downloader/Book1.xlsx"
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook.active

# Start row for data (e.g., row 2)
start_row = 2

driver = webdriver.Chrome()

# Read username and password from the text file
with open("credentials.txt", "r") as file:
    lines = file.readlines()
    username = lines[0].strip()
    password = lines[1].strip()

GBSS = "https://target.cbxportal.com/merchantWeb/login.action"

driver.get(GBSS)

# Find the username and password input fields and login button
username_field = driver.find_element("name", "j_username")  # Using the name attribute
password_field = driver.find_element("name", "j_password")  # Replace with the actual name
login_button = driver.find_element("class name", "submitButton")  # Replace with the actual class name

# Fill in the login credentials and press Enter to submit (assuming pressing Enter triggers login)
username_field.send_keys(username)
password_field.send_keys(password)
password_field.send_keys(Keys.ENTER)

# Assuming the button text is "QM"
qm_button = driver.find_element("xpath", "//span[text()='QM']")
qm_button.click()
time.sleep(25)

# Wait for the link to be clickable
wait = WebDriverWait(driver, 25)
link = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'Sample & Document Submission')]")))

# Click the link
link.click()

# # Click the link using JavaScript
# driver.execute_script("arguments[0].click();", link)

# Loop through rows and download files
while True:
    cell_value = sheet.cell(row=start_row, column=2).value
    if cell_value is None:
        break
    print(cell_value)
    # Wait for the page to load (you might need to adjust the wait time)
    time.sleep(5)
    driver.implicitly_wait(30)

    text_input = driver.find_element("css selector", "input#v_styleNo")

    # Clear any existing text and then input your desired text
    text_input.clear()
    text_input.send_keys(cell_value)

    search_button = driver.find_element("css selector", "button#btn_searchAdv")

    # Click the search button
    search_button.click()

    # Wait for the row to be clickable
    wait = WebDriverWait(driver, 10)
    row = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='x-grid3-cell-inner x-grid3-col-styleNo'][contains(text(), cell_value)]")))

    # Click the row
    row.click()

    # Wait for the details page to load (you might need to adjust the wait time)
    time.sleep(2)
    driver.implicitly_wait(30)
    # wait = WebDriverWait(driver, 10)

    print_tech_pack_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='btn_printBuyBrief']/span")))

    # Click the "Print Tech Pack" button
    print_tech_pack_button.click()

    # Wait for the printing process to initiate (you might need to adjust the wait time)
    time.sleep(2)

    # Switch to the new window
    new_window_handle = driver.window_handles[-1]  # Assuming the new window is the last one opened
    driver.switch_to.window(new_window_handle)

    # Find and select the radio button
    radio_button = driver.find_element("xpath", "//input[@id='optionChoice$d' and @value='allSample']")
    radio_button.click()

    # Find and click the "Print Tech Pack" button in the new window
    print_tech_pack_button_in_new_window = driver.find_element("xpath", "//*[@id='btn_printTechPack']/span")
    print_tech_pack_button_in_new_window.click()

    # Wait for the printing process to initiate (you might need to adjust the wait time)
    time.sleep(40)
    wait = WebDriverWait(driver, 30)

    # Switch back to the original window if necessary
    # driver.switch_to.window(driver.window_handles[0])

    # Find and click the "Search" button
    search_button = driver.find_element("xpath", "//*[@id='btn_search']/span")
    search_button.click()

    # Move to the next row
    start_row += 1

while True:
    pass
