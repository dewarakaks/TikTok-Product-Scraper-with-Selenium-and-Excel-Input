import time
import random
import pandas as pd
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Selenium configuration
def setup_driver():
    options = Options()
    options.headless = False
    options.add_argument('--disable-gpu')
    options.add_argument('--log-level=3')
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument(
        'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    )
    return webdriver.Chrome(options=options)

# Read URLs from Excel
def read_urls_from_excel(file_path):
    try:
        all_sheets = pd.read_excel(file_path, sheet_name=None)
        urls = []
        for sheet_name, df_sheet in all_sheets.items():
            if 'URL' in df_sheet.columns and 'ARTICLE' in df_sheet.columns:
                for _, row in df_sheet.iterrows():
                    if not pd.isna(row['URL']) and not pd.isna(row['ARTICLE']):
                        urls.append({
                            'URL': row['URL'],
                            'Article': row['ARTICLE']
                        })
            else:
                print(f"Sheet '{sheet_name}' missing required columns.")
        if not urls:
            print("No valid data found in the Excel file.")
            exit()
        return urls
    except FileNotFoundError:
        print("Error: Excel file not found.")
        exit()
    except Exception as e:
        print(f"Excel read error: {e}")
        exit()

# Parse sold info
def parse_sold_info(text):
    if not text:
        return None
    text = text.upper().replace(',', '').strip()
    match = re.search(r'([\d\.]+)(K?)', text)
    if not match:
        return None
    number, suffix = match.groups()
    number = float(number)
    if suffix == 'K':
        number *= 1000
    return int(number)

# Scrape data
def scrape_data(driver, url_data):
    try:
        driver.get(url_data['URL'])
        time.sleep(random.uniform(10, 15))

        # Extract title
        try:
            title = driver.find_element(By.CLASS_NAME, 'index-title--AnTxK').text
        except Exception:
            try:
                title = driver.find_element(By.CLASS_NAME, 'title-v0v6fK').text
            except Exception:
                title = "Title not found"

        # Extract price
        try:
            price = driver.find_element(By.CLASS_NAME, 'index-price--hHzq8').text
            if '-' in price:
                price_cleaned = int(re.sub(r'[^\d]', '', price.split('-')[0].strip()))
            else:
                price_cleaned = int(re.sub(r'[^\d]', '', price.strip()))
        except Exception:
            try:
                price = driver.find_element(By.CLASS_NAME, 'price-w1xvrw').text
                if '-' in price:
                    price_cleaned = int(re.sub(r'[^\d]', '', price.split('-')[0].strip()))
                else:
                    price_cleaned = int(re.sub(r'[^\d]', '', price.strip()))
            except Exception:
                price_cleaned = None

        # Extract and clean sold info
        try:
            sold_info = driver.find_element(By.CLASS_NAME, 'index-info__sold--Lgz8w').text
            sold_cleaned = parse_sold_info(sold_info)
        except Exception:
            try:
                sold_info = driver.find_element(By.CLASS_NAME, 'info__sold-ZdTfzQ').text
                sold_cleaned = parse_sold_info(sold_info)
            except Exception:
                sold_cleaned = None

        return {
            'URL': url_data['URL'],
            'Article': url_data['Article'],
            'Title': title,
            'Price': price_cleaned,
            'Total Sold': sold_cleaned,
            'Current Date': datetime.now()
        }

    except Exception as e:
        print(f"Error scraping {url_data['URL']}: {e}")
        return None

# Main
if __name__ == "__main__":
    try:
        print("Please select your Excel file containing the URLs...")
        Tk().withdraw()
        excel_file = askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not excel_file:
            print("No file selected. Exiting.")
            exit()

        urls = read_urls_from_excel(excel_file)
        driver = setup_driver()

        products = []
        for idx, url_data in enumerate(urls, start=1):
            print(f"Scraping {idx}/{len(urls)}")
            data = scrape_data(driver, url_data)
            if data:
                products.append(data)

        if products:
            output_file = f'tiktok_{datetime.now().strftime("%d %b %y")}.xlsx'
            pd.DataFrame(products).to_excel(output_file, index=False)
            print(f"\nData saved to: {output_file}")
        else:
            print("No data was scraped.")

    except Exception as e:
        print(f"Unexpected error: {e}")

    finally:
        if 'driver' in locals():
            driver.quit()