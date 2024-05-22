import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os
from concurrent.futures import ThreadPoolExecutor, as_completed

# Define the form URL
form_url = input("Input Google Form URL here: ")

# Load the Excel file
current_dir = os.path.dirname(os.path.realpath(__file__))
excel_file_path = os.path.join(current_dir, 'data.xlsx')
df = pd.read_excel(excel_file_path)

# Define the function to fill the form
def fill_form(row):
    driver = webdriver.Chrome()

    def fill_form_element(driver, input_xpath, textarea_xpath, value):
        try:
            element = driver.find_element(By.XPATH, input_xpath)
            element.send_keys(value)
        except Exception:
            try:
                element = driver.find_element(By.XPATH, textarea_xpath)
                element.send_keys(value)
            except Exception as e:
                print(f"Error locating or interacting with input/textarea element: {e}")

    try:
        driver.get(form_url)

        # Wait until the form is loaded
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, 'form'))
        )

        # Fill in the fields
        fill_form_element(driver, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input',
                          '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div[2]/textarea', row['name'])
        fill_form_element(driver, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input',
                          '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/textarea', str(row['ic']))
        fill_form_element(driver, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input',
                          '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div[2]/textarea', row['consultant name'])
        fill_form_element(driver, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input',
                          '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div[2]/textarea', str(row['consultant id']))
        fill_form_element(driver, '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div/div[1]/input',
                          '//*[@id="mG61Hd"]/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div[1]/div[2]/textarea', str(row['consultant phone']))

        # Select the "BOOKING METHOD"
        booking_method_xpath = '//*[@id="i25"]' if row['booking method'] == 'CASH BUYER' else '//*[@id="i28"]'
        driver.find_element(By.XPATH, booking_method_xpath).click()

        # Submit the form
        submit_button = driver.find_element(By.XPATH, '//*[@id="mG61Hd"]/div[2]/div/div[3]/div[1]/div[1]/div/span/span')
        submit_button.click()

        # Optionally, wait for a success message or redirection
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Your response has been recorded.")]'))
        )

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        driver.quit()

# Function to manage the submission of tasks to the executor
def process_rows(executor, rows):
    rows = rows.to_dict(orient='records')
    futures = {executor.submit(fill_form, row): row for row in rows[:num_workers]}
    rows = rows[num_workers:]

    while futures:
        for future in as_completed(futures):
            row = futures.pop(future)
            try:
                future.result()
            except Exception as e:
                print(f"An error occurred with row {row}: {e}")

            if rows:
                next_row = rows.pop(0)
                futures[executor.submit(fill_form, next_row)] = next_row

num_workers = 10  # Adjust this number based on your system's capability
with ThreadPoolExecutor(max_workers=num_workers) as executor:
    process_rows(executor, df)