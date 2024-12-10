import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoAlertPresentException
import test_utilities

# Initialize WebDriver
driver = webdriver.Chrome()
driver.get("https://only-testing-blog.blogspot.com/2014/05/form.html")
driver.maximize_window()

# Path to the Excel file
path = r"C:\\SeleniumPractice\\DataDrivenTesting.xlsx"

# Get the row count from the Excel sheet
rows = test_utilities.getRowCount(path, 'Sheet1')

# Iterate through the rows to fill the form
for r in range(2, rows + 1):
    first_name = test_utilities.readData(path, "Sheet1", r, 1)
    last_name = test_utilities.readData(path, "Sheet1", r, 2)
    email = test_utilities.readData(path, "Sheet1", r, 3)
    phone = test_utilities.readData(path, "Sheet1", r, 4)
    company = test_utilities.readData(path, "Sheet1", r, 5)

    print(f"Processing row {r}: {first_name}, {last_name}, {email}, {phone}, {company}")

    # Fill the form fields
    driver.find_element(By.NAME, "FirstName").clear()
    driver.find_element(By.NAME, "FirstName").send_keys(first_name)

    driver.find_element(By.NAME, "LastName").clear()
    driver.find_element(By.NAME, "LastName").send_keys(last_name)

    driver.find_element(By.NAME, "EmailID").clear()
    driver.find_element(By.NAME, "EmailID").send_keys(email)

    driver.find_element(By.NAME, "MobNo").clear()
    driver.find_element(By.NAME, "MobNo").send_keys(phone)

    driver.find_element(By.NAME, "Company").clear()
    driver.find_element(By.NAME, "Company").send_keys(company)  # Pass company value

    # Submit the form
    driver.find_element(By.XPATH, '//*[@id="post-body-8228718889842861683"]/div[1]/form/input[6]').click()
    time.sleep(5)  # Wait for alert to appear

    try:
        # Switch to the alert and accept it (click OK)
        alert = driver.switch_to.alert
        alert.accept()
        print(f"Alert handled for row {r}")

        # Check if the title is correct and write the result to the Excel file
        if driver.title == "Only Testing: Form":
            print("Test is Passed")
            test_utilities.writeData(path, "Sheet1", r, 6, "Passed")
        else:
            print("Test is Failed")
            test_utilities.writeData(path, "Sheet1", r, 6, "Failed")

    except NoAlertPresentException:
        print(f"No alert present for row {r}")
        test_utilities.writeData(path, "Sheet1", r, 6, "Failed")

    # Navigate back to the form page
    driver.back()
    time.sleep(2)

# Close the browser
driver.quit()
