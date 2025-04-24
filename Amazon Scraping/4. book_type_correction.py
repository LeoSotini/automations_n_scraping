from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import time
import random
import pandas as pd
import logging
import json

random_timer = random.randint(3,15)
options = Options()
options.debugger_address = "localhost:8989"
driver = webdriver.Edge(options=options)

feminism_df = pd.read_excel(r'C:\Users\leoso\Projects\Amazon Scraping\Livros Amazon.xlsx', sheet_name='Feminism_Kindle_Only')
antifeminism_df = pd.read_excel(r'C:\Users\leoso\Projects\Amazon Scraping\Livros Amazon.xlsx', sheet_name='AntiFeminism_Kindle_Only')
feminism_media_type_dict = dict(zip(feminism_df['Nome'], feminism_df['Link']))
antifeminism_media_type_dict = dict(zip(antifeminism_df['Nome'], antifeminism_df['Link']))

df_dict = {'Nome': [],
           'Link': [],
           'Fonte': [],
           'Divs': []}

def store_variables(dictionary):
    with open(r'C:\Users\leoso\Projects\data.txt', 'w') as file:
        json.dump(dictionary, file)

for col, values in feminism_media_type_dict.items():
    df_dict['Nome'].append(col)
    df_dict['Link'].append(values)
    df_dict['Fonte'].append('Feminismo')

for col, values in antifeminism_media_type_dict.items():
    df_dict['Nome'].append(col)
    df_dict['Link'].append(values)
    df_dict['Fonte'].append('Antifeminismo')


for i in range(len(df_dict['Nome'])):
    nome = df_dict['Nome'][i]
    link = df_dict['Link'][i]
    
    driver.get(link)
    time.sleep(random_timer)
    
    try:
        tmm_swatches_list = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'tmmSwatchesList'))
        )
    except TimeoutException:
        logging.error(f"{nome} - Timeout: Price slot not found.")
        df_dict['Divs'].append(0)
        continue

    try:
        divs = tmm_swatches_list.find_elements(
            By.CSS_SELECTOR, "div.a-row.formatsRow.a-ws-row"
        )
        qty = len(divs)
        df_dict['Divs'].append(qty)
        store_variables(df_dict)
        print(f'{nome} --- Done!')
    except NoSuchElementException:
        logging.error(f"{nome} - Price types or values not found.")
        df_dict['Divs'].append(0)
        continue