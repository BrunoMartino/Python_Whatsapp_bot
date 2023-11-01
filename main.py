#!/usr/bin/env python
# coding: utf-8

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from datetime import time
import schedule
import time as tm
import pandas as pd
import os
import pyperclip

service = Service(ChromeDriverManager().install())
browser = webdriver.Chrome(service = service)

browser.get('https://web.whatsapp.com/')
while len(browser.find_elements(By.ID, 'side')) < 1 :
  tm.sleep(1)
tm.sleep(2)  

table = pd.read_excel('./sheets/Grupos_De_Whatsapp.xlsx')

def load_table():
  global table
  new_table = pd.read_excel('./sheets/Grupos_De_Whatsapp.xlsx')
  table = new_table
  print(table)
  return table
load_table()

def send_message():
  for line in table.index:
    group = table.loc[line, "Grupos"]
    message = table.loc[line, "Mensagem"] 
    archive = table.loc[line, "Arquivo"] 
    status = table.loc[line, 'Status']

    archive_path = os.path.abspath(f'imgs/{archive}')
  
    def simple_send():    
      pyperclip.copy(message)
      browser.find_element('xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.CONTROL + 'v')
      tm.sleep(3)
      browser.find_element('xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]/p').send_keys(Keys.ENTER)
      tm.sleep(3)
      return

    def archive_send():
      browser.find_element('xpath','//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div').click()
      tm.sleep(2)
      browser.find_element('xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div/input').send_keys(archive_path)
      tm.sleep(5)
      pyperclip.copy(message)
      browser.find_element('xpath', '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p').send_keys(Keys.CONTROL + 'v')
      tm.sleep(2)
      browser.find_element('xpath', '//*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]/p').send_keys(Keys.ENTER)
      tm.sleep(2)
      return

    if status == "Ativo":
      browser.find_element('xpath', '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p').click()
      tm.sleep(1)
      browser.find_element('xpath', '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p').send_keys(group)
      tm.sleep(2)
      browser.find_element('xpath', '//*[@id="side"]/div[1]/div/div[2]/div[2]/div/div[1]/p').send_keys(Keys.ENTER)
      tm.sleep(3)

      if pd.isna(archive):
        simple_send()
      else:
        archive_send()  

time_1 = '10:00'
time_2 = '19:00'

schedule.every(180).seconds.do(load_table)
schedule.every().day.at(time_1).do(send_message)
schedule.every().day.at(time_2).do(send_message)

while True:
  schedule.run_pending()
  tm.sleep(1)       

