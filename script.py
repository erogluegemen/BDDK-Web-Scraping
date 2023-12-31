# Başlangıçta yüklü olmayan, eksik kütüphanelerin yüklenmesi.
!pip install selenium
!pip install dateparser
!pip install XlsxWriter

# Kodun çalışması için gerekli olan kütüphane importları.
import os
import time
import shutil
import datetime
import dateparser
import xlsxwriter

import numpy as np
import pandas as pd

from typing import Union, Optional

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service

# Uyarı mesajlarını önlemek için konfigürasyon ayarları.
import warnings
warnings.filterwarnings("ignore")

# Çalışma dizini (CWD) ve geçici ve ana veri dizinlerini tanımlayın.
CWD = os.getcwd()
temp_directory = "temp_bddk_files"
directory = 'bddk_files'
dir_list = [directory, temp_directory]

# Dizinlerin mevcut olup olmadığını kontrol edin, yoksa oluşturun.
for dir in dir_list:
    if not os.path.exists(dir):
        os.makedirs(dir)

# Selenium için tarayıcı seçeneklerini ve ayarlarını yapılandırın.
PATH = f'{CWD}/{temp_directory}'
chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : PATH}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--headless=new')
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36")


def get_date_from_website() -> list:
  '''
  BDDK web sitesinden tarihleri almak için kullanılan fonksiyon.
  '''

  # BDDK web sitesi URL'si
  URL = 'https://www.bddk.org.tr/bultenhaftalik'

  # Selenium için hizmet oluşturun ve tarayıcıyı başlatın.
  service = Service()
  driver = webdriver.Chrome(service=service, options=chrome_options)

  # Sayfanın yüklenmesi için bir süre bekleyin.
  time.sleep(3)

  # BDDK haftalık bülten sayfasına gidin.
  driver.get(URL)

  # Sayfanın tam yüklenmesi için bir süre daha bekleyin.
  time.sleep(3)

  # Tarihleri depolamak için bir liste oluşturun.
  date_list = []

  # Güncel son 2 tarihi almak için döngü oluşturun. [t, t-1]
  # (1 -> En güncel tarih,   2 -> En güncel tarihten bir önceki tarih)
  for i in range(1,3):
    # Web sitesinden tarihi alın ve düzenleyin.
    date_str = '-'.join(map(str,[driver.find_element("xpath",'//*[@id="Yil"]/option[1]').text] + driver.find_element("xpath",f'//*[@id="Donem"]/option[{i}]').text.split(' ')[0].split('/')))
    date = dateparser.parse(date_str)
    date = date.strftime('%Y-%m-%d')
    date_list.append(date)

  # Tarih listesini döndürün.
  return date_list


def seperate_folders(keyword:str) -> None:
  '''
  Belirli anahtar kelimeye sahip dosyalar ilgili klasöre taşımak için kullanılan fonksiyon.
  '''

  for filename in os.listdir(PATH):
    # Yapılacak işlemleri sadece dizindeki '.xls' uzantılı dosyalar için yap.
    if filename.endswith(".xls"):
      folder_name = keyword
      destination_folder = os.path.join(PATH, folder_name)

      # Anahtar kelimeye ait dosya yok ise oluştur.
      if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

      source_path = os.path.join(PATH, filename)
      destination_path = os.path.join(destination_folder, filename)

      # Dosyayı kaynak klasörden hedef klasöre taşı.
      shutil.move(source_path, destination_path)


def scrape_data(i) -> None:
    '''
    Verileri BDDK web sitesinden çekmek için kullanılan fonksiyon.
    '''

    # Taraf ve Tablo verilerini içeren sözlükler.
    # (Sayıların karşılık geldiği değerleri bulmak için BDDK'nın sitesinden sağ tık + incele yapabilirsiniz.)
    tarafDict = {
                10001:500, 10002:501, 10008:502, 10009:503, 10010:504,
                10003:505, 10004:506, 10005:507, 10006:508, 10007:511
                }

    tabloDict = {
                280:'Krediler', 281:'Takipteki Alacaklar', 282:'Menkul Değerler',
                283:'Mevduat', 284:'Diğer Bilanço Kalemleri', 285:'Bilanço Dışı İşlemler',
                }

    # 512 ve 513 numaralı bölümler raporda kullanılmayacağı için çıkarılmıştır.
    # Gereklilik dahilinde yorum satırından çıkarılarak tekrar eklenebilir.
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

    # Sayfanın yüklenmesi için bir süre bekleyin.
    time.sleep(3)

    # BDDK haftalık bülten sayfasına gidin.
    driver.get(URL)

    # Sayfanın tam yüklenmesi için bir süre daha bekleyin.
    time.sleep(3)

    # Tıklanacak kutuların ve butonların tanımlanması.
    date_box = '//*[@id="Donem"]'
    date_dropdown = f'//*[@id="Donem"]/option[{i}]'
    button_getir = '//*[@id="DonemTablosu"]/tbody/tr/td/form/div[4]/button'

    # Tarih seçme kutusunu açın.
    driver.find_element(By.XPATH, date_box).click()
    time.sleep(3)

    # İlgili tarih dönemini seçin.
    driver.find_element(By.XPATH, date_dropdown).click()
    time.sleep(3)

    # Getir butonuna tıklayın.
    driver.find_element(By.XPATH, button_getir).click()
    time.sleep(3)

    # Taraf kısmı baz alınarak, her bir taraf kutusu için sol taraftaki bilgi kutularını tek tek dön.
    # Örnek: Sektör-Krediler, Sektör-Takipteki Alacaklar .. Mevduat-Krediler, Mevduat-Takipteki Alacaklar
    for tar_k, tar_v in tarafDict.items():
        driver.find_element(By.XPATH, '//*[@id="tarafListesiItem-'+str(tar_k)+'"]').click()
        time.sleep(3)
        driver.find_element(By.XPATH, '//*[@id="BTN"]').click()

        for tab_k, tab_v in tabloDict.items():
            driver.find_element(By.XPATH, '//*[@id="tabloListesiItem-'+str(tab_k)+'"]').click()
            time.sleep(3)
            driver.find_element(By.XPATH, '//*[@id="BTN"]').click()
            time.sleep(1)

        # Çekilen verileri anahtar kelimelerini baz alarak farkı dosyalara yönlendir.
        seperate_folders(tarafNameDict[tar_v])

    time.sleep(2)
    driver.close()


def remove_unnecessary_files() -> None:
  '''
  Belirtilen dizindeki tüm dosyaları silmek için kullanılan fonksiyon.
  '''
  items = os.listdir(PATH)
  for item in items:
    item_path = os.path.join(PATH, item)
    if os.path.isdir(item_path):
      pass
    else:
      os.remove(item_path)


def rename_xls_files(directory:str=PATH) -> None:
    '''
     Belirtilen dizindeki tüm .xls dosyalarının isimlerini,
     klasör adlarına göre yeniden anlandırmak için kullanılan fonksiyon.
    '''

    # İsimlerini düzeltirken aynı zamanda istediğimiz sıraya göre sıralamasını istiyoruz.
    # Bu yüzden isimlendirmelerin başına 1-, 2- gibi ön ekleri ekliyoruz.
    folder_mapping = {
      'Krediler': '1-Krediler',
      'TakiptekiAlacaklar': '2-TakiptekiAlacaklar',
      'MenkulDeğerler': '3-MenkulDeğerler',
      'Mevduat': '4-Mevduat',
      'DiğerBilan&#231;oKalemleri': '5-DiğerBilançoKalemleri',
      'Bilan&#231;oDışıİşlemler': '6-BilançoDışıİşlemler',
      }

    # İsimlendirme kısmını yapan döngü.
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
  '''
  Belirtilen dizindeki tüm klasörleri ve alt klasörlerindeki tüm dosyaları
  bir sözlüğe aktarmak için kullanılan fonksiyon.
  '''

  # Dizindeki tüm klasörleri listeler.
  folders = os.listdir(PATH)

  # Her bir klasör için, dosyalar sözlüğe aktarılır.
  # Dosyaları tersten sıralayarak, en yeni dosyalar en başta gelir.
  ordered_dict = {}
  for folder in folders:
    # Her bir klasör için, o klasördeki tüm dosyaların bir listesini alır.
    temp_ordered_list = []
    for files in os.listdir(os.path.join(PATH,folder)):
      temp_ordered_list.append(files)
      temp_ordered_list = sorted(temp_ordered_list, reverse=False)
    ordered_dict[folder] = temp_ordered_list

  return ordered_dict


def clean_dict(dictionary:dict) -> dict:
  '''
  Belirtilen sözlükteki her bir klasör için, o klasördeki tüm dosyalardan,
  Haftalık Bülten numaralarını ayıklayıp tekrardan bir sözlüğe aktarır.
  '''

  # Temizlenmiş dosya listesini sözlüğe aktarır.
  cleaned_data = {}

  # Sözlükteki her bir klasör için, o klasördeki tüm dosyaların bir listesini alır.
  for key, paths in dictionary.items():
      # Her bir dosya için, haftalık bülten numarasını bulur.
      seen_numbers = set()
      cleaned_paths = []
      for path in paths:

          # Haftalık bülten numarasını bulur.
          start_index = path.find("HaftalikBulten _") + len("HaftalikBulten _")
          end_index = path.find("-", start_index)
          if end_index == -1:
              end_index = path.find("_", start_index)

          number = path[start_index:end_index]
          # Haftalık bülten numarası daha önce görülmediyse, dosyayı temizlenmiş dosya listesine ekler.
          if number not in seen_numbers:
              seen_numbers.add(number)
              cleaned_paths.append(path)

      if cleaned_paths:
        cleaned_data[key] = cleaned_paths

  return cleaned_data


def create_dataframes_dict(cleaned_dict:dict) -> dict:
  '''
  Belirtilen sözlükteki her bir klasör için, o klasördeki tüm .xls dosyalarından
  bir dataframe oluşturmak için kullanılan fonksiyon.
  '''

  # Sözlükteki her bir klasör için, o klasördeki tüm .xls dosyalarını okur.
  df_dict = {}

  for k, v in cleaned_dict.items():
      dfs = []
      for file_path in v:
          # .xls dosyasını okur.
          full_file_path = os.path.join(PATH, k, file_path)
          data = pd.read_html(full_file_path, encoding='utf8') # thousands='.'
          df = pd.concat(data, ignore_index=True)
          # Bu kısımdaki kontrol yukarıda yorum satırına aldığımız 512 ve 513 numaralı bölümler için.
          # Güncel rapor formatında onları almadığımız için bu koşula aslında gerek yok fakat yine de
          # ilerde alınması durumuna karşın koşul eklenmiştir. Kodun akışında herhangi bir etkisi yoktur.
          if len(df.columns) == 5:
              columns_with_5 = ['Index', 'Birim: Milyon TL', 'TP', 'YP', 'TOPLAM']
              df.columns = columns_with_5
              df = df.iloc[1:]
              df = df.drop(columns='Index')
          # 512/513 bölümleri olmadığı sürece hiçbir zaman else koşuluna girmeyecek.
          else:
              columns_with_3 = ['Index', 'Birim: Milyon TL', 'TOPLAM']
              df.columns = columns_with_3
              df['TP'] = 0.0
              df['YP'] = 0.0
              df = df[['Index', 'Birim: Milyon TL', 'TP', 'YP', 'TOPLAM']]
              df = df.iloc[1:]
              df = df.drop(columns='Index')
          dfs.append(df)

      # DataFrame'leri bir sözlüğe aktarır.
      if dfs:
          df_dict[k] = pd.concat(dfs, ignore_index=True)

  result_dfs = {}
  for k, df in df_dict.items():
      result_dfs[k] = df.copy()

  return result_dfs


def create_excel(result_dfs:dict) -> None:
  '''
  Belirtilen sözlükteki tüm Dataframe'leri bir Excel dosyasına
  kaydetmek için kullanılan fonksiyon.
  '''

  # Dosya adını ve yolunu ayarlar.
  file_name = 'data.xlsx'
  output_path = os.path.join(CWD,directory,file_name)

  # Dataframe'leri Excel dosyasına yazar.
  with pd.ExcelWriter(output_path) as writer:
    for k, df in result_dfs.items():
      df.to_excel(writer, sheet_name=k, index=False)


def reformat_excel(date:str) -> None:
  '''
  Belirtilen tarihe ait Excel dosyasın yeniden biçimlendirmesi için kullanılan fonksiyon.
  '''

  # Kalem kolonun değerlerini oluşturmak için 'Bilgi' kısmındaki değerler,
  # ve bu değerlerin taraf kısmında karşılık gelen sayılarının tanımlanması.
  krediler = ['Krediler'] * 22
  takipteki_alacaklar = ['Takipteki Alacaklar'] * 12
  menkul_degerler = ['Menkul Değerler'] * 13
  mevduat = ['Mevduat'] * 11
  mevduat_2 = ['Mevduat'] * 12
  diger_bilanco_kalemleri = ['Diğer Bilanço Kalemleri'] * 16
  bilanco_disi_islemler = ['Bilanço Dışı İşlemler'] * 4

  # Sektör tarafı hariç kalem değişkenini kullanıyoruz. Bütün kalemlerin toplamını belirtir.
  kalem = krediler + takipteki_alacaklar + menkul_degerler + mevduat + diger_bilanco_kalemleri + bilanco_disi_islemler
  # Kalem 2 istisnai bir durumdur. Sadece sektör tarafının mevudat kısmında 11 olması gereken değer 12'dir.
  kalem_2 = krediler + takipteki_alacaklar + menkul_degerler + mevduat_2 + diger_bilanco_kalemleri + bilanco_disi_islemler

  # Örnek üzerinden ilerlemek gerekirse, 22 adet 'Krediler' satırının olduğunu var sayarsak,
  # Krediler 101'den başlayıp n+1 formatında n+22 ye kadar ilerleyecek.
  # Bu sistem bütün Bilgi değerleri için geçerlidir. Hepsinin başlangıç değeri ve gittiği uzunluk farklıdır.
  start_values = {
      'Krediler': 101,
      'Takipteki Alacaklar': 201,
      'Menkul Değerler': 301,
      'Mevduat': 401,
      'Diğer Bilanço Kalemleri': 501,
      'Bilanço Dışı İşlemler': 601,
  }

  # data.xlsx dosyası geçici olarak oluşturduğumuz bir dosyadır.
  excel_file_path = os.path.join(CWD, directory, 'data.xlsx')
  excel_data = pd.ExcelFile(excel_file_path)

  # data.xlsx Excel'inin içindeki sheetler alınır.
  sheet_names = excel_data.sheet_names

  sheet_dataframes = {}

  # Her bir sheet teker teker dönülür. Gerekli kolonlar oluşturulur/eklenir.
  for sheet_name in sheet_names:
      sheet_df = excel_data.parse(sheet_name)
      sheet_df.insert(0, 'Taraf', sheet_name)

      # Yukarıda belirttiğimiz üzere sektör bu durumda istisna belirtmektedir.
      if sheet_name == 'Sektör':
        sheet_df.insert(1, 'Kalem', pd.Series(kalem_2))
        sheet_df.insert(0, 'Metrik',  sheet_df.groupby('Kalem').cumcount() + sheet_df['Kalem'].map(start_values))

      # Sektör harici kısımlar için normal işlemler ile devam edilir.
      else:
        sheet_df.insert(1, 'Kalem', pd.Series(kalem))
        sheet_df.insert(0, 'Metrik',  sheet_df.groupby('Kalem').cumcount() + sheet_df['Kalem'].map(start_values))

      # END_DT kolonun değerleri sabit tek bir değer alır.(döngü başına)
      # Bu değer o kolona sahip bütün satırlara ayn değeri yazar.
      sheet_df['END_DT'] =  date
      sheet_dataframes[sheet_name] = sheet_df

  # Tarih bazlı 2 adet excel çıktısı alıyoruz.
  # (Güncel Tarih, Güncel Tarih - 1)
  file_name = f'{date}.xlsx'
  output_excel_path = os.path.join(CWD, directory, file_name)
  with pd.ExcelWriter(output_excel_path) as writer:
      for sheet_name, sheet_df in sheet_dataframes.items():
          sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


def create_mevduatozel_sheet(date:str) -> None:
  '''
  Mevduat Özel var olan bir Taraf olmadığı ama raporda istendiği için
  manuel olarak oluşturuyoruz.
  Mevduat Özel = Mevduat Yabancı + Mevduat Yerli Özel formülü ile oluşuyor.
  '''
  # Dinamik bir şekilde tarihe özel 2 excel oluşturulur.
  file_name = f'{date}.xlsx'
  excel_path = os.path.join(CWD, directory, file_name)

  # Formüle dayanarak kullanılacak 2 Taraf'ın excel sheetleri çekilir.
  df_mevduat_yabanci = pd.read_excel(excel_path, sheet_name='MevduatYabancı')
  df_mevduat_yerliozel = pd.read_excel(excel_path, sheet_name='MevduatYerliÖzel')

  def clean_and_convert(value:str) -> float:
    '''
    Amerika-Avrupa arasındaki ondalık sayı sistemi farklılığın çözmek için
    kullanılan fonksiyon.
    '''

    # Fonksiyona gelen ifadedeki noktaları boşluğa, virgülleri noktaya çevirir.
    cleaned_value = str(value).replace('.', '').replace(',', '.')
    return float(cleaned_value)

  # Yukarıda yazdığımız fonksiyon 2 Excel'de de olan 3 sayısal kolona uygulanır.
  df_mevduat_yabanci['TP'] = df_mevduat_yabanci['TP'].apply(clean_and_convert)
  df_mevduat_yabanci['YP'] = df_mevduat_yabanci['YP'].apply(clean_and_convert)
  df_mevduat_yabanci['TOPLAM'] = df_mevduat_yabanci['TOPLAM'].apply(clean_and_convert)

  df_mevduat_yerliozel['TP'] = df_mevduat_yerliozel['TP'].apply(clean_and_convert)
  df_mevduat_yerliozel['YP'] = df_mevduat_yerliozel['YP'].apply(clean_and_convert)
  df_mevduat_yerliozel['TOPLAM'] = df_mevduat_yerliozel['TOPLAM'].apply(clean_and_convert)

  # Yeni, boş bir dataframe oluşturulur.
  df_mevduat_ozel = pd.DataFrame(columns=df_mevduat_yabanci.columns)

  # 2 Tarafın sayısal kolonlarının toplamı ile yeni kolonlar oluşturulur.
  df_mevduat_ozel['TP'] = df_mevduat_yabanci['TP'] + df_mevduat_yerliozel['TP']
  df_mevduat_ozel['YP'] = df_mevduat_yabanci['YP'] + df_mevduat_yerliozel['YP']
  df_mevduat_ozel['TOPLAM'] = df_mevduat_yabanci['TOPLAM'] + df_mevduat_yerliozel['TOPLAM']

  # Eksik kalan sayısal olmayan kolon değerleri herhangi bir Tarafın sheet'i ile doldurulabilir.
  df_mevduat_ozel['Metrik'] = df_mevduat_yabanci['Metrik']
  df_mevduat_ozel['Taraf'] = 'MevduatÖzel'
  df_mevduat_ozel['Kalem'] = df_mevduat_yabanci['Kalem']
  df_mevduat_ozel['Birim: Milyon TL'] = df_mevduat_yabanci['Birim: Milyon TL']
  df_mevduat_ozel['END_DT'] = df_mevduat_yabanci['END_DT']

  # Yeni oluşturulan dataframe'i Excel'e eklerken koyulacak sheet'in adı belirlenir.
  # Aynı Excel'in üzerine, append komutu ile yazılır.
  sheet_name = 'MevduatÖzel'
  with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
      df_mevduat_ozel.to_excel(writer, sheet_name = sheet_name, index=False)


def create_ozel_sheet(date:str) -> None:
  '''
  Özel var olan bir Taraf olmadığı ama raporda istendiği için
  manuel olarak oluşturuyoruz.
  Özel = Yabancı + Yerli Özel formülü ile oluşuyor.
  '''

  # Dinamik bir şekilde tarihe özel 2 excel oluşturulur.
  file_name = f'{date}.xlsx'
  excel_path = os.path.join(CWD, directory, file_name)

  # Formüle dayanarak kullanılacak 2 Taraf'ın excel sheetleri çekilir.
  df_yabanci = pd.read_excel(excel_path, sheet_name='Yabancı')
  df_yerliozel = pd.read_excel(excel_path, sheet_name='YerliÖzel')

  def clean_and_convert(value):
    '''
    Amerika-Avrupa arasındaki ondalık sayı sistemi farklılığın çözmek için
    kullanılan fonksiyon.
    '''

    # Fonksiyona gelen ifadedeki noktaları boşluğa, virgülleri noktaya çevirir.
    cleaned_value = str(value).replace('.', '').replace(',', '.')
    return float(cleaned_value)

  # Yukarıda yazdığımız fonksiyon 2 Excel'de de olan 3 sayısal kolona uygulanır.
  df_yabanci['TP'] = df_yabanci['TP'].apply(clean_and_convert)
  df_yabanci['YP'] = df_yabanci['YP'].apply(clean_and_convert)
  df_yabanci['TOPLAM'] = df_yabanci['TOPLAM'].apply(clean_and_convert)

  df_yerliozel['TP'] = df_yerliozel['TP'].apply(clean_and_convert)
  df_yerliozel['YP'] = df_yerliozel['YP'].apply(clean_and_convert)
  df_yerliozel['TOPLAM'] = df_yerliozel['TOPLAM'].apply(clean_and_convert)

  # Yeni, boş bir dataframe oluşturulur.
  df_ozel = pd.DataFrame(columns=df_yabanci.columns)

  # 2 Tarafın sayısal kolonlarının toplamı ile yeni kolonlar oluşturulur.
  df_ozel['TP'] = df_yabanci['TP'] + df_yerliozel['TP']
  df_ozel['YP'] = df_yabanci['YP'] + df_yerliozel['YP']
  df_ozel['TOPLAM'] = df_yabanci['TOPLAM'] + df_yerliozel['TOPLAM']

  # Eksik kalan sayısal olmayan kolon değerleri herhangi bir Tarafın sheet'i ile doldurulabilir.
  df_ozel['Metrik'] = df_yabanci['Metrik']
  df_ozel['Taraf'] = 'Özel'
  df_ozel['Kalem'] = df_yabanci['Kalem']
  df_ozel['Birim: Milyon TL'] = df_yabanci['Birim: Milyon TL']
  df_ozel['END_DT'] = df_yabanci['END_DT']

  # Yeni oluşturulan dataframe'i Excel'e eklerken koyulacak sheet'in adı belirlenir.
  # Aynı Excel'in üzerine, append komutu ile yazılır.
  sheet_name = 'Özel'
  with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a') as writer:
      df_ozel.to_excel(writer, sheet_name = sheet_name, index=False)


def transform_numeric_columns(date:str) -> None:
  '''
  Bütün sheetler tam olarak oluştuktan sonra, son Excel üzerindeki tüm sayısal kolonlara
  tranform işleminin uygulanmasını sağlayan fonksiyon.
  '''

  def clean_and_convert(value):
    '''
    Amerika-Avrupa arasındaki ondalık sayı sistemi farklılığın çözmek için
    kullanılan fonksiyon.
    '''

    # Fonksiyona gelen ifadedeki noktaları boşluğa, virgülleri noktaya çevirir.
    cleaned_value = str(value).replace('.', '').replace(',', '.')
    return float(cleaned_value)


  # Dinamik olarak tarih bazlı tekil Excel'in okunur.
  excel_file_path = os.path.join(CWD, directory, f'{date}.xlsx')
  excel_data = pd.ExcelFile(excel_file_path)

  # Bütün sheet'ler alınır.
  sheet_names = excel_data.sheet_names
  sheet_dataframes = {}

  # Bütün sheetlerde tek tek dolanılır.
  # Bulunan her sayısal kolona transform işlemi uygulanır.
  for sheet_name in sheet_names:
    sheet_df = excel_data.parse(sheet_name)
    sheet_df['TP'] = sheet_df['TP'].apply(clean_and_convert)
    sheet_df['YP'] = sheet_df['YP'].apply(clean_and_convert)
    sheet_df['TOPLAM'] = sheet_df['TOPLAM'].apply(clean_and_convert)
    sheet_dataframes[sheet_name] = sheet_df

  # Yeni oluşturulan dataframe'i Excel'e eklerken koyulacak sheet'in adı belirlenir.
  # Aynı Excel'in üzerine, append komutu ile yazılır.
  with pd.ExcelWriter(excel_file_path) as writer:
    for sheet_name, sheet_df in sheet_dataframes.items():
      sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


def create_weekly_report(date_list:list):
  '''
  Belirtilen tarihlere ait haftalık raporları oluşturur.
  '''

  # Parametre olarak alınan tarih listesinin ilk elemanı güncel tarih, ikinci elemanı güncel tarih - 1 formatındadır.
  current_date, previous_date = date_list

  # Çekilen tarih bilgileri, dosya isimlerini bulmak için kullanılr.
  filename_1, filename_2 = f'{current_date}.xlsx', f'{previous_date}.xlsx'

  # Dosya isimleri bulunduktan sonra bu dosyaların dizinleri tanımlanır.
  file_1, file_2 = os.path.join(CWD,directory,filename_1), os.path.join(CWD,directory,filename_2)

  # 2 Excel ayr ayrı okunur.
  current_week_dataframes = pd.read_excel(file_2, sheet_name=None)
  previous_week_dataframes = pd.read_excel(file_1, sheet_name=None)

  # Raporda olması gereken eksik kolonlar eklenir.
  columns_to_add = ['TP', 'YP', 'TOPLAM', 'END_DT']

  # Güncel dataframe verileri bir sözlük yapısında tutulur.
  updated_dataframes = {}

  # Dataframe bir döngü doğrultusunda güncellenir.
  for sheet_name in previous_week_dataframes.keys():
    previous_week_df = previous_week_dataframes[sheet_name]
    current_week_df = current_week_dataframes[sheet_name]

    current_week_selected = current_week_df[columns_to_add]

    updated_df = pd.concat([previous_week_df, current_week_selected], axis=1)

    updated_dataframes[sheet_name] = updated_df

  # Yeni oluşturulan dataframe'i Excel'e eklerken koyulacak sheet'in adı belirlenir.
  file_name = 'Sektör_Haftalık_Data.xlsx'
  with pd.ExcelWriter(os.path.join(CWD,directory,file_name)) as writer:
      for sheet_name, updated_df in updated_dataframes.items():
          updated_df.to_excel(writer, sheet_name=sheet_name, index=False)


def concat_excel_sheets() -> pd.DataFrame:
  '''
  'Sektör_Haftalık_Data.xlsx' dosyasındaki tüm sayfaları birleştirmek için kullanılan fonksiyon.
  '''

  # 'Sektör_Haftalık_Data.xlsx' dosyasının yolu belirtilir.
  excel_file_path = os.path.join(CWD, directory, 'Sektör_Haftalık_Data.xlsx')

  # Excel dosyası okunur.
  excel_data = pd.ExcelFile(excel_file_path)

  # Sheet isimleri alınır.
  sheet_names = excel_data.sheet_names

  # Her bir sheet için bir dataframe oluşturulur.
  sheet_dataframes = {}
  for sheet_name in sheet_names:
    sheet_df = excel_data.parse(sheet_name)
    sheet_dataframes[sheet_name] = sheet_df

  # Aşağıda verilen liste sırası doğrultusunda oluşturulan dataframeler sıralanır.
  order_list = ['Sektör', 'Mevduat', 'Kamu', 'YerliÖzel', 'Yabancı', 'Özel', 'MevduatKamu',
                'MevduatYabancı', 'MevduatYerliÖzel', 'Katılım', 'KalkınmaVeYatırım', 'MevduatÖzel']
  reordered_dataframes = {k: sheet_dataframes[k] for k in order_list}

  # Sıralanan dataframeler birleştirilip tek bir dataframe oluşturulur.
  merged_df = pd.concat(reordered_dataframes.values(), axis=0)
  return merged_df


def merge_columns(df:pd.DataFrame) -> pd.DataFrame:
  '''
  Dataframe'deki sütünlar birleştirmek için kullanılan fonksiyon.

  -> İlk 4 kolon sabit değerlerin olduğu kolonlardır.
  -> 4-8 arasındaki kolonlar güncel tarihe ait dinamik değerli kolonlardır.
  -> 8-12 arasındaki kolonlar güncel tarih - 1'e ait dinamik değerli kolonlardır.

  => 4-8 arasındaki değerler ilk 4 kolonu tutarak, 8-12 kolonlarının altına eklenecektir.
  => Son halinde sütün sayısında azalış, satır sayısında artış gözlenmelidir.
  '''

  # Sabit kolonlar ve değerleri belirlenir.
  constant_data = df.iloc[:,:4]

  # Güncel tarihe ait kolonlar ve dinamik değerleri belirlenir.
  first_part = df.iloc[:,4:8]
  first_data = pd.merge(constant_data.reset_index(drop=True),
                      first_part.reset_index(drop=True),
                      left_index=True,
                      right_index=True)

  # Güncel tarihe - 1'e ait kolonlar ve dinamik değerleri belirlenir.
  second_part = df.iloc[:,8:12]
  second_data = pd.merge(constant_data.reset_index(drop=True),
                      second_part.reset_index(drop=True),
                      left_index=True,
                      right_index=True)
  second_data = second_data.rename(columns={'TP.1':'TP', 'YP.1':'YP',
                                            'TOPLAM.1':'TOPLAM', 'END_DT.1':'END_DT'})

  # Dataframe'ler birleştirilir.
  df = pd.concat([second_data, first_data])
  return df


def get_dates_from_excel(df:pd.DataFrame) -> list:
  '''
  Excel dosyasından END_DT kolonuna ait tarih bilgilerini almak için kullanılan fonksiyon.
  '''

  # Tarih kolonu datetime formatına çevirilir.
  df['END_DT'] = pd.to_datetime(df['END_DT'], errors='coerce')

  # Datetime formatına çevirildikten sonra Ay/Gün/Yıl formatına dönüştürülür.
  df['END_DT'] = df['END_DT'].dt.strftime('%m/%d/%Y')

  # Unique tarih bilgileri listeye atılır.
  dates = df['END_DT'].unique().tolist()

  # Bu liste döndürülür.
  return dates


def create_report_data() -> pd.DataFrame:
  '''
  İki haftalık raporun verilerini birleştirmek ve
  istenilen rapor formatına getirmek için kullanılan fonksiyon.
  '''

  df_list = []
  for i in range(2):
    # Excel dosyasını açar ve veriyi dataframe'e alır.
    df = concat_excel_sheets()
    df = merge_columns(df)

    # Tarihleri alır.
    dates = get_dates_from_excel(df)

    # Tarihe göre filtreleme yapıp dataframe'i oluşturur.
    dataframe = df[df['END_DT'] == dates[i]]

    # Sabit kolonlar ve tarih kolonu belirlenir.
    constant_data = dataframe.iloc[:,:4]
    date_column = dataframe.iloc[:,-1:]

    # TP kısmını oluşturan merge işlemleri yapılır.
    TP_part = dataframe.iloc[:,4]
    TP_data = pd.merge(constant_data.reset_index(drop=True),
                          TP_part.reset_index(drop=True),
                          left_index=True,
                          right_index=True)
    TP = pd.merge(TP_data.reset_index(drop=True),
                          date_column.reset_index(drop=True),
                          left_index=True,
                          right_index=True)

    # YP kısmını oluşturan merge işlemleri yapılır.
    YP_part = dataframe.iloc[:,5:6]
    YP_data = pd.merge(constant_data.reset_index(drop=True),
                          YP_part.reset_index(drop=True),
                          left_index=True,
                          right_index=True)
    YP = pd.merge(YP_data.reset_index(drop=True),
                          date_column.reset_index(drop=True),
                          left_index=True,
                          right_index=True)

    # TOPLAM kısmını oluşturan merge işlemleri yapılır.
    TOPLAM_part = dataframe.iloc[:,6:7]
    TOPLAM_data = pd.merge(constant_data.reset_index(drop=True),
                          TOPLAM_part.reset_index(drop=True),
                          left_index=True,
                          right_index=True)
    TOPLAM = pd.merge(TOPLAM_data.reset_index(drop=True),
                          date_column.reset_index(drop=True),
                          left_index=True,
                          right_index=True)


    # Metrikler birleştirilmek üzere yeniden adlandırılır.
    TP = TP.rename(columns={'TP':'METRIK_DEGER'})
    YP = YP.rename(columns={'YP':'METRIK_DEGER'})
    TOPLAM = TOPLAM.rename(columns={'TOPLAM':'METRIK_DEGER'})

    # Döngü bazında geçici bir dataframe üzerinde birleştirme yapılır.
    df_temp = pd.concat([TP, YP, TOPLAM])
    df_list.append(df_temp)

  # İki dataframe birleştirilir.
  final_df = pd.concat(df_list)

  # Sonuç olarak final dataframe'i döndürülür.
  return final_df


def reformat_report_data(final_df:pd.DataFrame) -> pd.DataFrame:
  '''
  Rapor verilerini istenilen formata getirmek ve
  eksik olan değerleri oluşturmak için kullanılan fonksiyon.
  '''

  # Tarafların denk geldiği sayısal karşılıklar sözlük olarak tutulur.
  tarafNameDict = {
                  'Sektör': 500,
                  'Mevduat': 501,
                  'Kamu': 502,
                  'YerliÖzel': 503,
                  'Yabancı': 504,
                  'Özel': 505,
                  'MevduatKamu': 506,
                  'MevduatYabancı': 507,
                  'MevduatYerliÖzel': 508,
                  'Katılım': 511,
                  'KalkınmaVeYatırım': 512,
                  'MevduatÖzel': 513,
                  }

  # Sözlükdeki değerlere göre map edilir.
  final_df['BNK_KRM_KOD'] = final_df['Taraf'].map(tarafNameDict)

  # Metrik Değer ve Metrik Izlm Değer kolonlarındaki değerler milyon TL cinsinde tutulur.
  final_df['METRIK_DEGER'] = final_df['METRIK_DEGER'] * 1_000_000
  final_df['METRIK_IZLM_DEGER'] = final_df['METRIK_DEGER']

  # İş Akış No kolonunun değeri sabit 26 olarak belirlenir.
  const_akis_no = 26
  final_df['IS_AKIS_NO'] = const_akis_no

  # T - Y - P değerlerinin sayısı sabit 937 olarak belirlenir.
  values_count = 937
  T_list = ['T'] * values_count
  Y_list = ['Y'] * values_count
  P_list = ['P'] * values_count

  # T - Y - P kodları tek bir kolonda tutulur. Değerler toplanıp TL_YP_KOD kolonuna yazılır.
  TL_YP_KOD = T_list + Y_list + P_list + T_list + Y_list + P_list
  final_df['TL_YP_KOD'] = TL_YP_KOD

  # Rapor formatında olmayan gereksiz kolonlar atılır.
  final_df.drop(columns=['Metrik', 'Kalem', 'Birim: Milyon TL'], axis=1, inplace=True)

  # Metrik numarası ayarlanır.
  def generate_sequence(start:int=9002, end:int=9077, step:int=1) -> None:
    '''
    90002'den başlayıp 9077'ye kadar devam edip, bittiğinde baştan başlayan
    bir seri oluşturmak için kullanılan fonksiyon.
    '''

    current_value = start
    while current_value <= end:
      yield current_value
      current_value += step

  sequence = generate_sequence()
  new_column = []
  for value in sequence:
    new_column.append(value)

  total_rows = len(final_df)
  METRIK_NO = new_column * (total_rows // len(new_column)) + new_column[:total_rows % len(new_column)]
  final_df['METRIK_NO'] = METRIK_NO

  # Kolonlar doğru sıralamaya getirilir.
  final_df = final_df[['METRIK_NO', 'END_DT', 'TL_YP_KOD', 'METRIK_DEGER',
                       'METRIK_IZLM_DEGER', 'BNK_KRM_KOD', 'IS_AKIS_NO']]

  # Final dataframe'i döndürülür.
  return final_df


def read_output_excel(output_excel:str) -> pd.DataFrame:
  '''
  En güncel, geçmiş tarihli raporu okur.
  '''

  # Excel dosyasının dataframe'e çevrilir.
  output_excel_df = pd.read_excel(output_excel)

  # END_DT kolonu datatime formatına çevrilir.
  output_excel_df['END_DT'] = pd.to_datetime(output_excel_df['END_DT'], errors='coerce')

  # Datetime formatındaki değerler Ay/Gün/Yıl formatına çevrilir.
  output_excel_df['END_DT'] = output_excel_df['END_DT'].dt.strftime('%m/%d/%Y')

  # Excel dosyası, dataframe olarak döndürülür.
  return output_excel_df


def get_dates_from_excel(df:pd.DataFrame) -> list:
  '''
  Dataframe'in END_DT sütunundaki benzersiz tarih değerlerini
  liste olarak döndürmek için kullanılan fonksiyon.
  '''

  # END_DT kolonu datatime formatına çevrilir.
  df['END_DT'] = pd.to_datetime(df['END_DT'], errors='coerce')

  # Datetime formatındaki değerler Ay/Gün/Yıl formatına çevrilir.
  df['END_DT'] = df['END_DT'].dt.strftime('%m/%d/%Y')

  # END_DT kolonundaki benzersiz değerler listesine çevrilir.
  dates = df['END_DT'].unique().tolist()

  # Liste formatında tarihler döndürülür.
  return dates


def control_date(report_df:pd.DataFrame, output_excel_df:pd.DataFrame) -> Optional[Union[str, None]]:
  '''
  En güncel, geçmiş tarihli rapordaki tarihler ile web sitesinden çekile verilerin tarihlerini karşılaştırır.
  '''

  latest_date = output_excel_df['END_DT'].values.tolist()[-1]
  dates = get_dates_from_excel(report_df)
  for date in dates:
    if date == latest_date:
      return date
    else:
      pass

# Bütün kodları ve fonksiyonları çalıştıran ana fonksiyon.
def run():
    # Websitesinden en güncel 2 tarih çekilir.
    date_list = get_date_from_website()

    # Çekilen en güncel 2 tarihe ait veriler/Excel'ler indirilir.
    for i, date in enumerate(date_list, start=1):
      print(f'{date} tarihli veri çekiliyor..')

      # Verileri çeken fonksiyon.
      scrape_data(i=i)

      # Gereksiz dosyaları silen fonksiyon.
      remove_unnecessary_files()

      # Dosyaları yeniden adlandıran fonksiyon.
      rename_xls_files()

      # Sıralı verileri sözlük yapısında tutan fonksiyon.
      ordered_dict = create_dict()

      # Temizlenmiş sıralı verileri sözlük yapısında tutan fonksiyon.
      cleaned_dict = clean_dict(dictionary=ordered_dict)

      # Temizlenmiş sıralı veriler sözlüğü ile dataframe oluşturan fonksiyon.
      result_dfs = create_dataframes_dict(cleaned_dict=cleaned_dict)

      # Excel dosyasını oluşturan fonksiyon.
      create_excel(result_dfs=result_dfs)

      # Excel dosyasının formatını düzenleyen fonksiyon.
      reformat_excel(date=date)

      # Excelde eksik olan Mevduat Özel sheetini oluşturan fonksiyon.
      create_mevduatozel_sheet(date=date)

      # Excelde eksik olan Özel sheetini oluşturan fonksiyon.
      create_ozel_sheet(date=date)

      # Sayısal kolonlara transform işlemi uygulayan fonksiyon.
      transform_numeric_columns(date=date)

      # Geçici oluşturulan dosyaların temizlenmesi.
      shutil.rmtree(PATH, ignore_errors=True)
      print(f'{date} tarihli veri başarıyla çekildi!')

    # Haftalık formatta raporları hazırlayan fonksiyon.
    create_weekly_report(date_list=date_list)

    # Geçici oluşturulan dosyaların temizlenmesi.
    shutil.rmtree(PATH, ignore_errors=True)

    # Tarih bazlı geçici dosyaların belirlenmesi.
    file_1 = f'{date_list[0]}.xlsx'
    file_2 = f'{date_list[1]}.xlsx'
    file_3 = 'data.xlsx'

    # Bu dosyaların, dizinlerinin belirlenmesi.
    file_1_path = os.path.join(CWD,directory,file_1)
    file_2_path = os.path.join(CWD,directory,file_2)
    file_3_path = os.path.join(CWD,directory,file_3)

    # Bu dosyaların kaldırılması.
    os.remove(file_1_path)
    os.remove(file_2_path)
    os.remove(file_3_path)

    # Rapor formatını oluşturan fonksiyon.
    final_df = create_report_data()

    # Raporun son halini oluşturan fonksiyon.
    report_df = reformat_report_data(final_df=final_df)

    # Güncel tarihli excelin ana dize eklendikten sonra dosya yolunun bulunup okunması.
    files_in_directory = os.listdir(CWD)
    excel_files = [file for file in files_in_directory if file.endswith('.xlsx')]
    excel_file = excel_files[0]
    output_excel_path = os.path.join(CWD, excel_file)
    output_excel_df = read_output_excel(output_excel=output_excel_path)

    # Tarihleri kontrol eden fonksiyon.
    date_result = control_date(report_df, output_excel_df)

    # Eğer eski rapordaki tarihler ile güncel çekilen tarihler arasında
    # hiç bir ortak tarih yok ise if condition'a girer.
    if date_result == None:
      # Yeni veriler eklenir.
      concatenated_df = pd.concat([output_excel_df, report_df])
      latest_date = concatenated_df['END_DT'].values.tolist()[-1]
      latest_date = datetime.datetime.strptime(latest_date, "%m/%d/%Y").strftime("%Y-%m-%d")

      # Yeni eklenen veriler güncel excelin üstüne yazılmayacağı için (override etmeyeceği için)
      # güncel tarihli yeni bir excel oluşturulur. Bu süreç her 2 haftalık veri geldiğinde tekrarlanır.
      concatenated_df.to_excel(f'{latest_date}.xlsx')

      # En son bddk_files dosyası da silinir.
      bddk_files_path = f'{CWD}/{directory}'
      shutil.rmtree(bddk_files_path, ignore_errors=True)

    # Eğer eski rapordaki tarihler ile güncel çekilen tarihler arasında
    # ortak bir tarih var ise herhangi bir veri güncellenmesi durumuna karşın
    # TRUNCATE/UPDATE yapılır.
    else:
      # Tarih var mı diye kontrol yapılır, var ise indexleri belirlenir.
      list_to_drop = output_excel_df[output_excel_df['END_DT']==date_result].index

      # Belirlenen tarihlere ait indexler droplanır.
      output_excel_df.drop(index=list_to_drop, inplace=True)

      # Yeni veriler eklenir.
      concatenated_df = pd.concat([output_excel_df, report_df])
      latest_date = concatenated_df['END_DT'].values.tolist()[-1]
      latest_date = datetime.datetime.strptime(latest_date, "%m/%d/%Y").strftime("%Y-%m-%d")

      # Yeni eklenen veriler güncel excelin üstüne yazılmayacağı için (override etmeyeceği için)
      # güncel tarihli yeni bir excel oluşturulur. Bu süreç her 2 haftalık veri geldiğinde tekrarlanır.
      concatenated_df.to_excel(f'{latest_date}.xlsx')

      # En son bddk_files dosyası da silinir.
      bddk_files_path = f'{CWD}/{directory}'
      shutil.rmtree(bddk_files_path, ignore_errors=True)

# run fonksiyonun çalıştırılması ile bütün süreç ayağa kaldırılır.
run()
