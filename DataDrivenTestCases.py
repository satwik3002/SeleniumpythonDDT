import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import test_utilities

# Initialize WebDriver
driver = webdriver.Chrome()
driver.get("https://www.letskodeit.com/login")
driver.maximize_window()

# Path to the Excel file
path = r"C:\SeleniumPractice\LoginTest (1).xlsx"

# Get the row count from the Excel sheet
rows = test_utilities.getRowCount(path, 'Sheet1')

# Iterate through the rows for login
for r in range(2, rows + 1):
    username = test_utilities.readData(path, "Sheet1", r, 1)  # Column 1 for username
    password = test_utilities.readData(path, "Sheet1", r, 2)  # Column 2 for password

    # Wait for the email field to be visible and interactable
    time.sleep(2)  # Wait for a few seconds to ensure the elements are loaded
    email_field = driver.find_element(By.NAME, 'email')
    email_field.clear()  # Clear any pre-filled text
    email_field.send_keys(username)

    # Wait for the password field to be interactable
    time.sleep(2)  # Allow some time for the password field to load
    password_field = driver.find_element(By.NAME, 'password')
    password_field.clear()  # Clear any pre-filled text
    password_field.send_keys(password)

    # Perform login
    driver.find_element(By.XPATH, "//button[@id='login']").click()
    time.sleep(5)

    # Perform logout
    time.sleep(2)  # Wait for the user menu to be visible
    user_menu = driver.find_element(By.XPATH, "//img[contains(@class,'zl-navbar-rhs-img')]")
    user_menu.click()
    time.sleep(2)  # Wait for the logout option to appear
    logout_option = driver.find_element(By.XPATH, "//a[contains(text(),'Logout')]")
    logout_option.click()
    time.sleep(5)

    # Check if the title is correct and write the result to the Excel file
    if driver.title == "Home Page":
        print("Test is Passed")
        test_utilities.writeData(path, "Sheet1", r, 3, "Passed")
    else:
        print("Test is Failed")
        test_utilities.writeData(path, "Sheet1", r, 3, "Failed")

    driver.find_element(By.XPATH, "//a[normalize-space()='Sign In']").click()
    time.sleep(5)


