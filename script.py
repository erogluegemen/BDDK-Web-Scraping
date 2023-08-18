# Libraries
import os
import time
import shutil
import numpy as np
import pandas as pd
from typing import Union
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service

# Create folders if not exists
CWD = os.getcwd()
temp_directory = "temp_bddk_files"
directory = 'bddk_files'
dir_list = [directory, temp_directory]
for dir in dir_list:
    if not os.path.exists(dir):
        os.makedirs(dir)

# Configurations
PATH = f'{CWD}/{temp_directory}'
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : PATH}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--headless=new') # comment this to debug


def scrape_data() -> None:
    print('Running...')

    tarafDict = {
                10001:500, 10002:501, 10008:502, 10009:503, 10010:504, 
                10003:505, 10004:506, 10005:507, 10006:508, 10007:511
                }

    tabloDict = {
                280:'Krediler', 281:'Takipteki Alacaklar', 282:'Menkul Değerler',
                283:'Mevduat', 284:'Diğer Bilanço Kalemleri', 285:'Bilanço Dışı İşlemler', 
                286:'Bankalarda Saklanan Menkul Değerler - 1', 287:'Bankalarda Saklanan Menkul Değerler - 2',
                288:'Yabancı Para Pozisyonu'
                }

    URL = 'https://www.bddk.org.tr/bultenhaftalik'
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    time.sleep(3)
    driver.get(URL)
    time.sleep(3)

    for k, v in tabloDict.items():
        driver.find_element(By.XPATH, '//*[@id="tabloListesiItem-'+str(k)+'"]').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="BTN"]').click()

        for t in tarafDict.keys():
            driver.find_element(By.XPATH, '//*[@id="tarafListesiItem-'+str(t)+'"]').click()
            time.sleep(3)
            driver.find_element(By.XPATH, '//*[@id="BTN"]').click()

    time.sleep(2)

    print('Completed!')
    driver.close()


def seperate_folders() -> None:
    folder_mapping = {
        'Krediler' : 'Krediler',
        'TakiptekiAlacaklar' : 'Takipteki Alacaklar',
        'MenkulDeğerler' : 'Menkul Değerler',
        'Mevduat' : 'Mevduat',
        'DiğerBilan&#231;oKalemleri' : 'Diğer Bilanço Kalemleri',
        'Bilan&#231;oDışıİşlemler' : 'Bilanço Dışı İşlemler',
        'BSMD1' : 'Bankalarda Saklanan Menkul Değerler - 1',
        'BSMD2' : 'Bankalarda Saklanan Menkul Değerler - 2',
        'YabancıParaPozisyonu' : 'Yabancı Para Pozisyonu',
                    }

    for filename in os.listdir(PATH):
        if filename.endswith(".xls"):
            keyword = filename.split("_")[1]
            folder_name = folder_mapping.get(keyword, "Other") 
            
            destination_folder = os.path.join(PATH, folder_name)
            
            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)
            
            source_path = os.path.join(PATH, filename)
            destination_path = os.path.join(destination_folder, filename)
            
            shutil.move(source_path, destination_path)


def extract_number(filename:str) -> Union[int, float]:
    parts = filename.split("(")
    if len(parts) > 1:
        number = parts[-1].split(")")[0]
        if number.isdigit():
            return int(number)
    return float('inf') 


def sort_files(sub_path:str) -> list:
    dir_list = os.listdir(PATH + '/' + sub_path)
    sorted_files = sorted(dir_list, key=extract_number)

    non_numerical_files = [file for file in sorted_files if not file.isdigit()]
    sorted_files = [file for file in sorted_files if file not in non_numerical_files]
    sorted_files = non_numerical_files + sorted_files

    #file = sorted_files.pop()
    #sorted_files.insert(0, file)
    return sorted_files


def create_excels() -> None:
    for folder in os.listdir(PATH):
        if folder == '.DS_Store' or folder == 'Yabancı Para Pozisyonu':
            continue

        sorted_files = sort_files(folder)

        temp_df_list = []

        for files in sorted_files:
            if folder == '.DS_Store':
                continue

            file_path = PATH + '/' + folder + '/' + files

            data = pd.read_html(file_path, encoding='utf-8')
            df = pd.concat(data, ignore_index=True)

            columns = ['Index', f'{folder.replace(" ", "_")}', 'TP', 'YP', 'TOPLAM']
            df.columns = columns

            df = df.iloc[1:]
            df = df.drop(columns='Index')

            temp_df_list.append(df)

        concatenate_df = pd.concat(temp_df_list)
        file_name = concatenate_df.columns[0]
        concatenate_df.to_excel(f'{directory}/{file_name}.xlsx')


def run():
    scrape_data()
    seperate_folders()
    create_excels()

run()
