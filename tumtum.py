from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException , NoSuchElementException
import pandas as pd
import time

# Initialize the WebDriver
browser = webdriver.Chrome()
browser.maximize_window()

# Open the Moneycontrol website
browser.get("https://www.moneycontrol.com")
time.sleep(20)

# Find and enter "Itc" in the search box
search_box = browser.find_element(By.XPATH, '//input[@id="search_str"]')
search_box.send_keys("Itc")
search_box.send_keys(Keys.ENTER)

# Wait for search results to load and click on the company link
time.sleep(5)
company = browser.find_element(By.CSS_SELECTOR, '#mc_mainWrapper > div.PA10 > div.PA10 > div > table > tbody > tr:nth-child(4) > td:nth-child(1) > p > a > strong')
company.click()

# Wait for the page to load
time.sleep(5)

# Directly navigate to the historical prices page using its URL
historical_prices_url = "https://www.moneycontrol.com/stocks/histstock.php?sc_id=ITC&mycomp=ITC"
browser.get(historical_prices_url)
time.sleep(5)

dropdown = browser.find_element(By.ID, "ex")

# Create a Select object
select = Select(dropdown)

# Select the "NSE" option by its value
select.select_by_visible_text("NSE")
time.sleep(5)

dropdown1 = browser.find_element(By.XPATH, '//select[@name="mth_frm_yr"]')
select = Select(dropdown1)
select.select_by_visible_text("2003")
time.sleep(5)

go_button = browser.find_element(By.XPATH, '(//input[@type="image"])[2]')
go_button.click()
time.sleep(5)

table = browser.find_element(By.XPATH, '(//table[@class="tblchart"])')
rows = table.find_elements(By.TAG_NAME, "tr")
data = []
for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")
    if len(cells) > 0:
        row_data = [cell.text for cell in cells]
        data.append(row_data)
df = pd.DataFrame(data, columns=["Date", "Open", "High", "Low", "Close"])
print(df)
excel_filename = "historical_data.xlsx"
df.to_excel(excel_filename, index=False)


print(f"Excel file '{excel_filename}' has been created.")

browser.quit()
