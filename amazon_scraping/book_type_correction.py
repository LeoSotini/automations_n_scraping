from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import logging
import pandas as pd

url = r'https://www.amazon.com.br/s?k=feminismo&i=stripbooks&s=exact-aware-popularity-rank&__mk_pt_BR=%C3%85M%C3%85%C5%BD%C3%95%C3%91&crid=1CEA01FML03B7&qid=1739483595&sprefix=feminismo%2Cstripbooks%2C182&ref=sr_st_exact-aware-popularity-rank&ds=v1%3AlQqJNF9gMMVAR8tTfp2S4Yal3%2BSrQSMQHltU5OQFEsU'
feminism_df = pd.DataFrame(pd.read_excel(r'C:\Users\leoso\Projects\Amazon Scraping\Livros Amazon.xlsx', sheet_name = 'Feminism_Kindle_Only'))
links_list = feminism_df['Link']
driver = webdriver.Edge()
driver.get(url)

tmm_swatches_list = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, 'tmmSwatchesList'))
)

divs = tmm_swatches_list.find_elements(
    By.CSS_SELECTOR, "div.a-row.formatsRow.a-ws-row"
    )

quantidade = len(divs)
print(f"Quantidade de divs encontradas: {quantidade}")

def capture_prices(driver, dictionary, url):
    driver.get(url)
    try:
        price_slot = WebDriverWait(driver, 10).until(
            WebDriverWait(driver, 10).until(((By.CSS_SELECTOR, 'div.a-row.formatsRow.a-ws-row'))
        ))
    except TimeoutException:
        logging.error(f"{dictionary} - Timeout: Price slot not found.")

    try:
        price_types = price_slot.find_elements(By.CLASS_NAME, 'slot-title')
        prices_values = price_slot.find_elements(By.CLASS_NAME, 'slot-price')
    except NoSuchElementException:
        logging.error(f"{dictionary} - Price types or values not found.")

    if len(price_types) != len(prices_values):
        logging.warning(f"{dictionary} - Mismatch in the number of price types and values.")
    for index, type_element in enumerate(price_types, start=1):
        try:
            dictionary[f'Tipo de livro {index}'] = type_element.text
            dictionary[f'Valor {index}'] = prices_values[index - 1].text
        except IndexError as e:
            logging.error(f"{dictionary} - Index error when processing prices: {e}")
            continue