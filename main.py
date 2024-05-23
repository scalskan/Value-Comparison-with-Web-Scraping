import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Initialize Chrome WebDriver
options = webdriver.ChromeOptions()
options.add_argument('--headless')  # Enable headless mode
driver = webdriver.Chrome(options=options)

# Read the Excel file
df = pd.read_excel('C:\\Users\\user\\Desktop\\WebScraper\\YourExcelFile.xlsx') #Your xlsx file

# Add price columns
df['FoundDiscountedPrice'] = ""
df['FoundOriginalPrice'] = ""

# Iterate over each product code
for index, row in df.iterrows():
    product_code = row['COLUIMN NAME'] #Excel column name to search in search bar.
    search_url = "YOURLINK" #Your link

    driver.get(search_url)
    time.sleep(3)  # Wait for the page to load

    try:
        # Find the search box and search for the product code
        search_box = driver.find_element(By.XPATH, "//input[@id='header-search-input']")
        search_box.send_keys(product_code)
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)  # Wait for search results to load

        # Find the first product in search results and go to the product page
        product_link = driver.find_element(By.CSS_SELECTOR, 'a[data-testid="product-tile"]')
        product_link.click()
        time.sleep(3)  # Wait for the product page to load

        # Get the discounted price
        discounted_price_element = driver.find_element(By.CSS_SELECTOR, 'span.sDq_FX._4sa1cA.dgII7d.Km7l2y')
        discounted_price = discounted_price_element.text.strip().replace('€', '').replace(',', '.')
        df.at[index, 'FoundDiscountedPrice'] = discounted_price

        # Get the original price
        original_price_element = driver.find_element(By.XPATH, "//p[@class='_0xLoFW u9KIT8 vSgP6A _7ckuOK']/span[2]")
        original_price = original_price_element.text.strip().replace('€', '').replace(',', '.')
        df.at[index, 'FoundOriginalPrice'] = original_price

    except Exception as e:
        df.at[index, 'FoundDiscountedPrice'] = 'Product Not Found'
        df.at[index, 'FoundOriginalPrice'] = 'Product Not Found'
        print(f"Error: {e}")

# Save the updated Excel file
df.to_excel('C:\\Users\\user\\Desktop\\WebScraper\\NEWFILE.xlsx', index=False) #New xlsx result file

# Quit the WebDriver
driver.quit()
