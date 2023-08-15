# Libraries
import os
import time
import locale
import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service

# Configurations
locale.setlocale(locale.LC_ALL, 'tr_TR')

chrome_options = webdriver.ChromeOptions()
prefs = {'download.default_directory' : f'{os.getcwd()}/bddk_files'}
chrome_options.add_experimental_option('prefs', prefs)
chrome_options.add_argument('--headless=new')

URL = 'https://www.bddk.org.tr/bultenhaftalik'

service = Service()

print('Running...')

driver = webdriver.Chrome(service=service, options=chrome_options)
time.sleep(3)
driver.get(URL)
time.sleep(3)

excel_button = driver.find_element(By.XPATH, '//*[@id="BTN"]').click()

time.sleep(3)

print('Completed!')
driver.close()
