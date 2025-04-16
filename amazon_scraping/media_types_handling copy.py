import pandas as pd
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import logging

# Set up basic logging configuration.
logging.basicConfig(level=logging.INFO)

def load_variables():
    # Make sure the file path is correct and accessible.
    with open(r'C:\Users\leoso\Projects\data.txt', 'r', encoding='utf-8') as file:
        return json.load(file)

# Load variables and create DataFrame.
data = load_variables()
df = pd.DataFrame(data)

# Filter and convert the DataFrame subsets to dictionaries.
df_feminism_correct = df[(df['Divs'] != 1) & (df['Fonte'] == 'Feminismo')].to_dict(orient='list')
df_antifeminism_correct = df[(df['Divs'] != 1) & (df['Fonte'] == 'Antifeminismo')].to_dict(orient='list')

def init_driver():
    options = Options()
    # Connect to an already running Edge debugger session if needed.
    options.debugger_address = "localhost:8989"
    # Add window size before instantiating the driver.
    options.add_argument("window-size=1920,1080")
    driver = webdriver.Edge(options=options)
    return driver

def capture_media_types_n_prices(driver):
    result = {}
    max_slots = 4

    try:
        # Wait until the element with ID 'tmmSwatchesList' is loaded.
        tmm_swatches_list = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'tmmSwatchesList'))
        )
    except TimeoutException:
        logging.error("Timeout: Media type slot not found for URL: %s", driver.current_url)
        for i in range(1, max_slots + 1):
            result[f'Tipo Livro {i}'] = "N/A"
            result[f'Preço {i}'] = "N/A"
        return result

    try:
        slot_titles = tmm_swatches_list.find_elements(By.CSS_SELECTOR, ".slot-title")
        slot_prices = tmm_swatches_list.find_elements(By.CSS_SELECTOR, ".slot-price")
        
        logging.info("Preços encontrados: %s", [sp.text for sp in slot_prices])
        
        # Extract up to max_slots data; if a slot is missing, return "N/A".
        for i in range(1, max_slots + 1):
            tipo_text = slot_titles[i - 1].text if i <= len(slot_titles) else "N/A"
            preco_text = slot_prices[i - 1].text if i <= len(slot_prices) else "N/A"
            result[f'Tipo Livro {i}'] = tipo_text
            result[f'Preço {i}'] = preco_text
        
        return result
    except NoSuchElementException:
        logging.error("Media types or prices not found on page: %s", driver.current_url)
    except Exception as e:
        logging.error("Unexpected error on page %s: %s", driver.current_url, e)
    
    # Default result for errors.
    for i in range(1, max_slots + 1):
        result[f'Tipo Livro {i}'] = "N/A"
        result[f'Preço {i}'] = "N/A"
    return result

def process_links(link_dict):
    results = {}
    # Iterate over the links from the provided dictionary.
    for url in link_dict.get('Link', []):
        driver = init_driver()
        try:
            driver.get(url)
            # Capture media types and prices.
            media_result = capture_media_types_n_prices(driver)
            results[url] = media_result
        except Exception as e:
            logging.error("Error processing %s: %s", url, e)
            results[url] = {"Error": str(e)}
        finally:
            driver.quit()
    return results

def store_dict_to_file(data, filename="data_2.json"):
    try:
        with open(filename, 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=4)
        logging.info("Data successfully saved to %s", filename)
    except Exception as e:
        logging.error("Error writing to file %s: %s", filename, e)

# Optional: If you want to open a homepage before processing, you can do so:
# driver = init_driver()
# driver.get('https://www.amazon.com.br/')
# driver.quit()

# Process the links from df_feminism_correct (or choose df_antifeminism_correct as needed).
results = process_links(df_antifeminism_correct)
store_dict_to_file(results)
print(results)