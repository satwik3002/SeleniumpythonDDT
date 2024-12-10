import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import test_utilities

# Initialize WebDriver
driver = webdriver.Firefox()
driver.get("https://demo.opencart.com.gr/")
driver.maximize_window()

# Verify the title of the page
expected_title = "Your Store"
actual_title = driver.title
if actual_title == expected_title:
    print("Title is correct: 'Your Store'")
else:
    print(f"Title is incorrect. Expected: '{expected_title}', but got: '{actual_title}'")
    driver.quit()  # Exit the script if the title is incorrect
    exit()

# Path to the Excel file
path = r"C:\\SeleniumPractice\\UserDetails.xlsx"

# Get the row count from the Excel sheet
rows = test_utilities.getRowCount(path, 'Sheet1')

# Iterate through the rows for account creation
for r in range(2, rows + 1):
    first_name = test_utilities.readData(path, "Sheet1", r, 1)  # Column 1: First Name
    last_name = test_utilities.readData(path, "Sheet1", r, 2)   # Column 2: Last Name
    email = test_utilities.readData(path, "Sheet1", r, 3)       # Column 3: Email
    telephone = test_utilities.readData(path, "Sheet1", r, 4)   # Column 4: Telephone
    password = test_utilities.readData(path, "Sheet1", r, 5)    # Column 5: Password
    confirm_password = test_utilities.readData(path, "Sheet1", r, 6)  # Column 6: Confirm Password

    # Navigate to Register page
    driver.find_element(By.XPATH, "//a[@title='My Account']").click()
    time.sleep(2)
    driver.find_element(By.XPATH, "//a[normalize-space()='Register']").click()
    time.sleep(3)

    # Fill in the registration form
    driver.find_element(By.ID, "input-firstname").clear()
    driver.find_element(By.ID, "input-firstname").send_keys(first_name)
    time.sleep(3)

    driver.find_element(By.ID, "input-lastname").clear()
    driver.find_element(By.ID, "input-lastname").send_keys(last_name)
    time.sleep(3)

    driver.find_element(By.ID, "input-email").clear()
    driver.find_element(By.ID, "input-email").send_keys(email)
    time.sleep(3)

    driver.find_element(By.ID, "input-telephone").clear()
    driver.find_element(By.ID, "input-telephone").send_keys(str(telephone))
    time.sleep(3)

    driver.find_element(By.ID, "input-password").clear()
    driver.find_element(By.ID, "input-password").send_keys(password)
    time.sleep(3)

    driver.find_element(By.ID, "input-confirm").clear()
    driver.find_element(By.ID, "input-confirm").send_keys(confirm_password)
    time.sleep(3)

    # Agree to Privacy Policy and Submit
    driver.find_element(By.NAME, "agree").click()
    driver.find_element(By.XPATH, "//input[@value='Continue']").click()
    time.sleep(5)

    # Verify Account Creation
    expected_success_title = "Your Account Has Been Created!"
    actual_success_title = driver.title
    if actual_success_title == expected_success_title:
        print(f"Test Passed for row {r}")
        test_utilities.writeData(path, "Sheet1", r, 7, "Passed")  # Column 7: Test Result
    else:
        print(f"Test Failed for row {r}")
        test_utilities.writeData(path, "Sheet1", r, 7, "Failed")

    # Return to Home Page
    driver.get("http://demo.opencart.com.gr/")
    time.sleep(3)

# Close the browser
driver.quit()
