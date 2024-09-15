from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time


def get_stock_quantity(ean):
    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-extensions")

    # Initialize WebDriver
    s = Service('D:/Python Files/WebScraping/chromedriver-win64/chromedriver-win64/chromedriver.exe')
    driver = webdriver.Chrome(service=s, options=chrome_options)

    try:
        # Force the language to English by appending `&language=en_GB`
        search_url = f"https://www.amazon.nl/s?k={ean}&language=en_GB"
        driver.get(search_url)

        # Handle cookie popup if present
        try:
            cookie_button = driver.find_element(By.ID, "sp-cc-accept")
            cookie_button.click()
            time.sleep(2)
        except:
            pass  # Continue if no cookie consent is present

        # Explicitly wait for the search results to load
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.s-main-slot div.s-result-item'))
            )
        except Exception as e:
            print(f"Error finding search results for EAN {ean}: {e}")
            return {
                'EAN': ean,
                'Stock Info': 'Error',
                'Product URL': '',
                'Title': '',
                'Price': '',
                'Rating': '',
                'Number of Items': '',
                'Item Volume': '',
                'ASIN': ''
            }

        # Locate the first product in search results
        product_links = driver.find_elements(By.CSS_SELECTOR, 'div.s-main-slot div.s-result-item h2 a')
        if not product_links:
            print(f"No products found for EAN {ean}")
            return {
                'EAN': ean,
                'Stock Info': 'Incorrect EAN',
                'Product URL': '',
                'Title': '',
                'Price': '',
                'Rating': '',
                'Number of Items': '',
                'Item Volume': '',
                'ASIN': ''
            }

        # Click on the first product link
        product_links[0].click()

        # Wait for the product page to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'productTitle'))
        )

        # Extract product information
        product_url = driver.current_url
        try:
            title = driver.find_element(By.ID, 'productTitle').text.strip()
        except:
            title = 'N/A'

        # Updated ASIN extraction
        try:
            asin = driver.find_element(By.XPATH, "//span[contains(text(), 'ASIN')]/following-sibling::span").text.strip()
        except:
            asin = 'N/A'

        # Updated Price extraction
        try:
            price = driver.find_element(By.CSS_SELECTOR, '.a-price .a-price-whole').text.strip()
            price_fraction = driver.find_element(By.CSS_SELECTOR, '.a-price .a-price-fraction').text.strip()
            price = price + ',' + price_fraction
        except:
            price = 'N/A'

        try:
            rating_element = driver.find_element(By.TAG_NAME, 'a')
            rating = rating_element.find_element(By.XPATH, "//span[@class='a-size-base a-color-base']").text.strip()
        except:
            rating = 'N/A'

        # Extract stock quantity or availability information
        try:
            stock_info_element = driver.find_element(By.ID, 'availability')
            stock_info = stock_info_element.text.strip()
            if not stock_info:
                stock_info = stock_info_element.find_element(By.XPATH,
                                                             "//span[@class='a-size-base a-color-price a-text-bold']").text.strip()
        except:
            stock_info = 'Not Available'

        # Extract number of items and item volume
        try:
            number_of_items = driver.find_element(By.XPATH, "//div[@id='productDetails_detailBullets_sections1']//th[contains(text(), 'Number of items') or contains(text(), 'Quantity')]/following-sibling::td").text.strip()
        except:
            number_of_items = 'N/A'

        try:
            item_volume = driver.find_element(By.XPATH, "//div[@id='productDetails_detailBullets_sections1']//th[contains(text(), 'Item volume') or contains(text(), 'Volume')]/following-sibling::td").text.strip()
        except:
            item_volume = 'N/A'

        return {
            'EAN': ean,
            'Stock Info': stock_info,
            'Product URL': product_url,
            'Title': title,
            'Price': price,
            'Rating': rating,
            'Number of Items': number_of_items,
            'Item Volume': item_volume,
            'ASIN': asin
        }

    except Exception as e:
        print(f"Error retrieving data for EAN {ean}: {e}")
        return {
            'EAN': ean,
            'Stock Info': 'Error',
            'Product URL': '',
            'Title': '',
            'Price': '',
            'Rating': '',
            'Number of Items': '',
            'Item Volume': '',
            'ASIN': ''
        }

    finally:
        # Quit the driver
        driver.quit()


# List of EANs to search
ean_list = [
    '4004675109941', '7615400039326', '7615400761340', '7615400190904',
    '7615400191307', '7615400191314', '7615400190652', '7615400190706',
    '7615400190836', '7615400190881','7615400193264', '7615400774869',
    '7615400774999', '7615400759613', '7615400193608', '4002632800658',
    '7613329007952', '4031101609676', '4031101609683', '4031101609690',
    '4031101609706', '4031101609720', '4031101609737', '4031101609744'
]

data = []

# Iterate over EANs and get product data
for ean in ean_list:
    product_info = get_stock_quantity(ean)
    print(product_info)
    data.append(product_info)

# Save to Excel
df = pd.DataFrame(data)
df.to_excel('product_stock_info_updated.xlsx', index=False)
