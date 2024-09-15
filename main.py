from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

def get_stock_quantity(ean):
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-extensions")

    s = Service('D:/Python Files/WebScraping/chromedriver-win64/chromedriver-win64/chromedriver.exe')
    driver = webdriver.Chrome(service=s, options=chrome_options)

    number_of_items = 'Not Available'
    item_volume = 'Not Available'
    item_weight = 'Not Available'

    try:
        search_url = f"https://www.amazon.nl/s?k={ean}&language=en_GB"
        driver.get(search_url)

        try:
            cookie_button = driver.find_element(By.ID, "sp-cc-accept")
            cookie_button.click()
            time.sleep(2)
        except:
            pass

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
                'Number of Items': number_of_items,
                'Item Volume': item_volume,
                'Item Weight': item_weight,
                'ASIN': ''
            }

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
                'Number of Items': number_of_items,
                'Item Volume': item_volume,
                'Item Weight': item_weight,
                'ASIN': ''
            }

        product_links[0].click()

        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'productTitle'))
        )

        product_url = driver.current_url
        try:
            title = driver.find_element(By.ID, 'productTitle').text.strip()
        except:
            title = 'N/A'

        try:
            asin = driver.find_element(By.XPATH, "//span[contains(text(), 'ASIN')]/following-sibling::span").text.strip()
        except:
            asin = 'N/A'

        try:
            price = driver.find_element(By.CSS_SELECTOR, '.a-price .a-price-whole').text.strip()
            price_fraction = driver.find_element(By.CSS_SELECTOR, '.a-price .a-price-fraction').text.strip()
            price = '€' + price + '.' + price_fraction
        except:
            price = 'N/A'

        try:
            rating_element = driver.find_element(By.TAG_NAME, 'a')
            rating = rating_element.find_element(By.XPATH, "//span[@class='a-size-base a-color-base']").text.strip()
        except:
            rating = 'N/A'

        try:
            stock_info_element = driver.find_element(By.ID, 'availability')
            stock_info = stock_info_element.text.strip()
            if not stock_info:
                stock_info = stock_info_element.find_element(By.XPATH,
                                                             "//span[@class='a-size-base a-color-price a-text-bold']").text.strip()
        except:
            stock_info = 'Not Available'

        try:
            table_rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
            for row in table_rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) == 2:
                    heading = cells[0].text
                    data = cells[1].text
                    if "Item volume" in heading:
                        item_volume = data
                    elif "Number of items" in heading:
                        number_of_items = data
                    elif "Item weight" in heading:
                        item_weight = data
        except:
            pass

        return {
            'EAN': ean,
            'Stock Info': stock_info,
            'Product URL': product_url,
            'Title': title,
            'Price': price,
            'Rating': rating,
            'Number of Items': number_of_items,
            'Item Volume': item_volume,
            'Item Weight': item_weight,
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
            'Number of Items': number_of_items,
            'Item Volume': item_volume,
            'Item Weight': item_weight,
            'ASIN': ''
        }

    finally:
        driver.quit()

ean_list = [
    '4004675109941', '7615400039326', '7615400761340', '7615400190904',
    '7615400191307', '7615400191314', '7615400190652', '4031101609744',
    '7615400190706', '7615400190836', '7615400190881', '7615400193264',
    '7615400774869', '7615400774999', '7615400759613', '7615400193608',
    '4002632800658', '7613329007952', '4031101609676', '4031101609683',
    '4031101609690', '4031101609706', '4031101609720', '4031101609737',
    ]

data = []

for ean in ean_list:
    product_info = get_stock_quantity(ean)
    print(product_info)
    data.append(product_info)

df = pd.DataFrame(data)
df.to_excel('product_info.xlsx', index=False)
