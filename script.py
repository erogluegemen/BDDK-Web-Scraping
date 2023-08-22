# -*- coding: utf-8 -*-
"""

!pip install selenium
!pip install dateparser

"""

# Libraries
import os
import time
import shutil
import dateparser
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
chrome_options.add_argument('--headless=new')
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36")


def get_date() -> str:
  URL = 'https://www.bddk.org.tr/bultenhaftalik'
  service = Service()
  driver = webdriver.Chrome(service=service, options=chrome_options)
  time.sleep(3)
  driver.get(URL)
  time.sleep(3)

  date_str = '-'.join(map(str,[driver.find_element("xpath",'//*[@id="Yil"]/option[1]').text] + driver.find_element("xpath",'//*[@id="Donem"]/option[1]').text.split(' ')[0].split('/')))
  date = dateparser.parse(date_str)
  date = date.strftime('%Y-%m-%d')
  return date


def seperate_folders(keyword:str) -> None:
    for filename in os.listdir(PATH):
        if filename.endswith(".xls"):
            folder_name = keyword
            destination_folder = os.path.join(PATH, folder_name)

            if not os.path.exists(destination_folder):
                os.makedirs(destination_folder)

            source_path = os.path.join(PATH, filename)
            destination_path = os.path.join(destination_folder, filename)

            shutil.move(source_path, destination_path)


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

    tarafNameDict = {
                    500:'Sektör',
                    501:'Mevduat',
                    502:'MevduatKamu',
                    503:'MevduatYabancı',
                    504:'MevduatYerliÖzel',
                    505:'KalkınmaVeYatırım',
                    506:'Katılım',
                    507:'Kamu',
                    508:'Yabancı',
                    511:'YerliÖzel',
                    #512:'KalkınmaYatırım',
                    #513:'MevduatÖzel',
                    }

    URL = 'https://www.bddk.org.tr/bultenhaftalik'
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    time.sleep(3)
    driver.get(URL)
    time.sleep(3)

    driver.find_element(By.XPATH, '//*[@id="Donem"]').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="Donem"]/option[3]').click()
    time.sleep(3)
    driver.find_element(By.XPATH, '//*[@id="DonemTablosu"]/tbody/tr/td/form/div[4]/button').click()
    time.sleep(3)

    for tar_k, tar_v in tarafDict.items():
        driver.find_element(By.XPATH, '//*[@id="tarafListesiItem-'+str(tar_k)+'"]').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="BTN"]').click()

        for tab_k, tab_v in tabloDict.items():
            driver.find_element(By.XPATH, '//*[@id="tabloListesiItem-'+str(tab_k)+'"]').click()
            time.sleep(3)
            driver.find_element(By.XPATH, '//*[@id="BTN"]').click()
            time.sleep(1)

        seperate_folders(tarafNameDict[tar_v])

    time.sleep(2)

    print('Completed!')
    driver.close()


def remove_unnecessary_files() -> None:
  items = os.listdir(PATH)
  for item in items:
      item_path = os.path.join(PATH, item)
      if os.path.isdir(item_path):
          pass
      else:
          os.remove(item_path)


def rename_xls_files(directory:str=PATH) -> None:
    folder_mapping = {
      'Krediler': '1-Krediler',
      'TakiptekiAlacaklar': '2-TakiptekiAlacaklar',
      'MenkulDeğerler': '3-MenkulDeğerler',
      'Mevduat': '4-Mevduat',
      'DiğerBilan&#231;oKalemleri': '5-DiğerBilançoKalemleri',
      'Bilan&#231;oDışıİşlemler': '6-BilançoDışıİşlemler',
      'BSMD1': '7-BankalardaSaklananMenkulDeğerler-1',
      'BSMD2': '8-BankalardaSaklananMenkulDeğerler-2',
      'YabancıParaPozisyonu': '9-YabancıParaPozisyonu',
      }

    for folder_name in os.listdir(directory):
        folder_path = os.path.join(directory, folder_name)
        if os.path.isdir(folder_path):
            for filename in os.listdir(folder_path):
                if filename.lower().endswith('.xls'):
                    for key, value in folder_mapping.items():
                        if key in filename:
                            new_filename = filename.replace(key, value)
                            file_path = os.path.join(folder_path, filename)
                            new_file_path = os.path.join(folder_path, new_filename)
                            os.rename(file_path, new_file_path)


def create_dict() -> dict:
  ordered_dict = {}
  for folder in os.listdir(PATH):
    temp_ordered_list = []
    for files in os.listdir(os.path.join(PATH,folder)):
      temp_ordered_list.append(files)
      temp_ordered_list = sorted(temp_ordered_list, reverse=False)
    ordered_dict[folder] = temp_ordered_list

  return ordered_dict


def clean_dict(dictionary:dict) -> dict:
  cleaned_data = {}

  for key, paths in dictionary.items():
      seen_numbers = set()

      cleaned_paths = []

      for path in paths:
          start_index = path.find("HaftalikBulten _") + len("HaftalikBulten _")
          end_index = path.find("-", start_index)
          if end_index == -1:
              end_index = path.find("_", start_index)

          number = path[start_index:end_index]

          if number not in seen_numbers:
              seen_numbers.add(number)
              cleaned_paths.append(path)

      if cleaned_paths:
        cleaned_data[key] = cleaned_paths

  return cleaned_data


def create_dataframes_dict(cleaned_dict:dict) -> dict:
  df_dict = {}

  for k, v in cleaned_dict.items():
      dfs = []
      for file_path in v:
          full_file_path = os.path.join(PATH, k, file_path)
          data = pd.read_html(full_file_path, encoding='utf8', thousands='.')
          df = pd.concat(data, ignore_index=True)
          if len(df.columns) == 5:
              columns_with_5 = ['Index', 'Birim: Milyon TL', 'TP', 'YP', 'TOPLAM']
              df.columns = columns_with_5
              df = df.iloc[1:]
              df = df.drop(columns='Index')
          else:
              columns_with_3 = ['Index', 'Birim: Milyon TL', 'TOPLAM']
              df.columns = columns_with_3
              df['TP'] = 0.0
              df['YP'] = 0.0
              df = df[['Index', 'Birim: Milyon TL', 'TP', 'YP', 'TOPLAM']]
              df = df.iloc[1:]
              df = df.drop(columns='Index')
          dfs.append(df)

      if dfs:
          df_dict[k] = pd.concat(dfs, ignore_index=True)

  result_dfs = {}
  for k, df in df_dict.items():
      result_dfs[k] = df.copy()

  return result_dfs


def create_excel(result_dfs:dict) -> None:
  file_name = 'data.xlsx'
  output_path = os.path.join(CWD,directory,file_name)
  with pd.ExcelWriter(output_path) as writer:
    for k, df in result_dfs.items():
      df.to_excel(writer, sheet_name=k, index=False)


def reformat_excel() -> None:
  krediler = ['Krediler'] * 22
  takipteki_alacaklar = ['Takipteki Alacaklar'] * 12
  menkul_degerler = ['Menkul Değerler'] * 13
  mevduat = ['Mevduat'] * 11
  diger_bilanco_kalemleri = ['Diğer Bilanço Kalemleri'] * 16
  bilanco_disi_islemler = ['Bilanço Dışı İşlemler'] * 4
  bsmd_1 = ['Bankalarda Saklanan Menkul Değerler - 1'] * 45
  bsmd_2 = ['Bankalarda Saklanan Menkul Değerler - 2'] * 45
  yabanci_para_pozisyonu = ['Yabancı Para Pozisyonu'] * 10

  kalem = krediler + takipteki_alacaklar + menkul_degerler + mevduat + diger_bilanco_kalemleri + bilanco_disi_islemler + bsmd_1 + bsmd_2 + yabanci_para_pozisyonu

  start_values = {
      'Krediler': 101,
      'Takipteki Alacaklar': 201,
      'Menkul Değerler': 301,
      'Mevduat': 401,
      'Diğer Bilanço Kalemleri': 501,
      'Bilanço Dışı İşlemler': 601,
      'Bankalarda Saklanan Menkul Değerler - 1': 701,
      'Bankalarda Saklanan Menkul Değerler - 2': 801,
      'Yabancı Para Pozisyonu': 901,
  }

  date = get_date()

  excel_file_path = os.path.join(CWD, directory, 'data.xlsx')
  excel_data = pd.ExcelFile(excel_file_path)

  sheet_names = excel_data.sheet_names

  sheet_dataframes = {}

  for sheet_name in sheet_names:
      sheet_df = excel_data.parse(sheet_name)
      sheet_df.insert(0, 'Taraf', sheet_name)
      sheet_df.insert(1, 'Kalem', pd.Series(kalem))
      sheet_df.insert(0, 'Metrik',  sheet_df.groupby('Kalem').cumcount() + sheet_df['Kalem'].map(start_values))
      sheet_df['END_DT'] =  date
      sheet_dataframes[sheet_name] = sheet_df

  file_name = f'{date}.xlsx'
  output_excel_path = os.path.join(CWD, directory, file_name)
  with pd.ExcelWriter(output_excel_path) as writer:
      for sheet_name, sheet_df in sheet_dataframes.items():
          sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


def run():
    scrape_data()
    remove_unnecessary_files()
    rename_xls_files()
    ordered_dict = create_dict()
    cleaned_dict = clean_dict(dictionary=ordered_dict)
    result_dfs = create_dataframes_dict(cleaned_dict=cleaned_dict)
    create_excel(result_dfs=result_dfs)
    reformat_excel()
    
run()