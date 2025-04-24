from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import logging
import time
import random
import json
import pandas as pd
import re

# Database Creation

#base_url = r'https://www.amazon.com.br/s?k=feminismo&i=stripbooks&s=exact-aware-popularity-rank&__mk_pt_BR=%C3%85M%C3%85%C5%BD%C3%95%C3%91&crid=1CEA01FML03B7&qid=1739483595&sprefix=feminismo%2Cstripbooks%2C182&ref=sr_st_exact-aware-popularity-rank&ds=v1%3AlQqJNF9gMMVAR8tTfp2S4Yal3%2BSrQSMQHltU5OQFEsU'
base_url = r'https://www.amazon.com.br'

random_timer = random.randint(3,15)
driver = webdriver.Edge()

books_dict = {}
selling_rank = 1
page = 1
page_ranking = 1

def open_driver(url):
    driver.get(url)
    time.sleep(random_timer)

open_driver(base_url)

# Database Creation

base_url = r'https://www.amazon.com.br/s?k=feminismo&i=stripbooks&s=exact-aware-popularity-rank&__mk_pt_BR=%C3%85M%C3%85%C5%BD%C3%95%C3%91&crid=1CEA01FML03B7&qid=1739483595&sprefix=feminismo%2Cstripbooks%2C182&ref=sr_st_exact-aware-popularity-rank&ds=v1%3AlQqJNF9gMMVAR8tTfp2S4Yal3%2BSrQSMQHltU5OQFEsU'

random_timer = random.randint(3,15)
driver = webdriver.Edge()

books_dict = {}
selling_rank = 1
page = 1
page_ranking = 1

def open_driver(url):
    driver.get(url)
    time.sleep(random_timer)

def get_book_links():
    global selling_rank, page_ranking

    try:
        link_html_slot = driver.find_element(
            By.CSS_SELECTOR,
            "div.s-main-slot.s-result-list.s-search-results.sg-row"
        )
    except Exception as e:
        print(f"Error finding main slot on page {page}: {e}")
        return

    links = link_html_slot.find_elements(
            By.CSS_SELECTOR,
            "a.a-link-normal.s-line-clamp-2.s-link-style.a-text-normal"
        )

    print(f"Found {len(links)} links on page {page}")

    for link in links:
        href = link.get_attribute("href")
        if 'sspa/click' not in href:
            books_dict[selling_rank] = {
                'Nome': link.text,
                'Overall Ranking': selling_rank,
                'Link': href,
                'Page': page,
                'Page Ranking': page_ranking
            }
            selling_rank += 1
            page_ranking += 1

def go_next_page():
    global page, page_ranking

    page_ranking = 1
    page += 1
    
    try:
        next_link_element = driver.find_element(By.CSS_SELECTOR, "a.s-pagination-item.s-pagination-next")
        next_page_url = next_link_element.get_attribute("href")
        if not next_page_url:
            raise Exception("Next page URL not found.")
        
        driver.get(next_page_url)

        WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "div.s-main-slot.s-result-list.s-search-results.sg-row")
                )
            )
        time.sleep(random_timer)
    
    except Exception as e:
        raise Exception(f"Error navigating to page {page}: {e}")

def store_variables():
    with open(r'C:\Users\leoso\Projects\data.txt', 'w') as file:
        json.dump(books_dict, file)

def load_variables():
    with open(r'C:\Users\leoso\Projects\data.txt', 'r') as file:
        recovered_dict = json.load(file)
        recovered_dict

def run_4_links():
    open_driver(base_url)

    while True:
        try:
            get_book_links()
            store_variables()
            go_next_page()
        except Exception as e:
            print(f'Error found: {e}')
            load_variables()
            driver.quit()
            break

# Data scraping

def capture_bottom_data(driver, dictionary):
    try:
        product_detail_slot = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#detailBulletsWrapper_feature_div'))
            )
    except TimeoutException:
       f'{dictionary} - Timeout: Product bottom data slot not found.'

    try:    
        product_detail_list = product_detail_slot.find_element(
            By.CSS_SELECTOR, 
            'ul.a-unordered-list.a-nostyle.a-vertical.a-spacing-none.detail-bullet-list'
            )
    except NoSuchElementException:
        logging.error(f'{dictionary} - Bottom data list not found')
        
    li_elements = product_detail_list.find_elements(
        By.TAG_NAME, 
        "li"
        )
    for li in li_elements:
        try:
            label_span = li.find_element(By.CSS_SELECTOR, "span.a-text-bold")
            label_text = label_span.text.replace(':','').strip()

            value_span = li.find_element(By.CSS_SELECTOR, "span.a-text-bold + span")
            value_text = value_span.text.strip()

            dictionary[label_text] = value_text
        except NoSuchElementException as e:
            logging.warning(f'{dictionary} - Missing element in list item: {e}')
            continue

def capture_ranking(driver, dictionary):
    try:
        ranking_element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="detailBulletsWrapper_feature_div"]/ul[1]/li'))
                )
    except TimeoutException:
        logging.error(f'{dictionary} - Timeout: Ranking data slot not found.')

    pattern = r'Nº\s*([\d\.]+)\s*em\s*([^\n(]+)'
    matches = re.findall(pattern, ranking_element.text)

    if not matches:
        logging.info(f'{dictionary} - No ranking matches found.')

    for i, (rank_str, category_str) in enumerate(matches, start=1):
        try:
            rank_int = int(rank_str.replace('.', ''))
        except ValueError as e:
            logging.warning(f'{dictionary} - Error converting rank to integer: {e}')
            rank_int = None
        
        category = category_str.strip()

        dictionary[f'Cat{i}'] = category
        dictionary[f'Rank_Cat{i}'] = rank_int

def capture_authors(driver, dictionary):
    try:
        authors_slot = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "bylineInfo"))
        )
    except TimeoutException:
        logging.error(f'{dictionary} - Timeout: Authors slot not found.')
    
    try:
        authors_spans = authors_slot.find_elements(
            By.CSS_SELECTOR, 
            "span.author.notFaded"
            )
    except NoSuchElementException:
        logging.error(f'{dictionary} - Authors spans not found.')

    count = 1

    for span in authors_spans:
        try:
            name_element = span.find_element(
                By.TAG_NAME, 
                "a"
                )
            name_text = name_element.text.strip()

            role_element = span.find_element(
                By.CSS_SELECTOR, 
                "span.a-color-secondary"
                )
            role_text = role_element.text.strip()

            if 'Autor' in role_text:
                dictionary[f"Autor {count}"] = name_text
                count += 1
        except NoSuchElementException as e:
            logging.warning(f"{dictionary} - Missing author element: {e}")
            continue

def capture_prices(driver, dictionary):
    try:
        price_slot = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.a-row.formatsRow.a-ws-row'))
        )
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

def capture_ratings(driver, dictionary):
    rating_slot = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="detailBulletsWrapper_feature_div"]/ul[2]'))
            )

    try:
        rating_slot
    except TimeoutException:
        logging.error(f"{dictionary} - Timeout: Ratings slot not found.")
        dictionary['Avaliação_média'] = None
        dictionary['Nº Avaliações'] = None
    else:
        try:
            rating_element = rating_slot.find_element(
                By.CSS_SELECTOR, 
                'span.a-size-base.a-color-base'
                )
            rating_text = rating_element.text.strip().replace(',', '.')
            dictionary['Avaliação_média'] = float(rating_text)
        except (NoSuchElementException, ValueError) as e:
            logging.warning(f"{dictionary} - Error retrieving average rating: {e}")
            dictionary['Avaliação_média'] = None

        try:
            num_ratings_element = rating_slot.find_element(
                By.CSS_SELECTOR, 
                '#acrCustomerReviewText'
                )
            match = re.search(r'(\d+(?:\.\d+)*)', num_ratings_element.text)
            if match:
                dictionary['Nº Avaliações'] = int(match.group(1).replace('.', ''))
            else:
                dictionary['Nº Avaliações'] = None
        except NoSuchElementException as e:
            logging.warning(f"{dictionary} - Number of ratings element not found: {e}")
            dictionary['Nº Avaliações'] = None

def store_data(path, dictionary):
    with open(path, 'w') as file:
        json.dump(dictionary, file)

with open(r'C:\Users\leoso\Projects\data.txt', 'r') as file:  
    books_dict = json.load(file)

storing_path = r'C:\Users\leoso\Projects\books_dict.txt'
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
random_timer = random.randint(3,15)
driver = webdriver.Edge()

for index, values in books_dict.items():
    link = values['Link']

    driver.get(values['Link'])
    time.sleep(random_timer)

    capture_bottom_data(driver = driver, dictionary = books_dict[index])
    capture_ratings(driver = driver, dictionary = books_dict[index])
    capture_authors(driver = driver, dictionary = books_dict[index])
    capture_prices(driver = driver, dictionary = books_dict[index])
    capture_ranking(driver = driver, dictionary = books_dict[index])
    store_data(path = storing_path, dictionary = books_dict)