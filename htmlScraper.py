# Import required modules
from lxml import html
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from openpyxl import load_workbook
import time
import glob
print(glob.glob("C:/Users/ACiurba/Desktop/similarwebdownloads/*"))
x_path_dict= {'A' : '//*[@id="overview"]/div/div/div/div[6]/div/table/tbody/tr[1]/td[2]',
'B' : '//*[@id="overview"]/div/div/div/div[1]/p[2]',
'C' : '//*[@id="overview"]/div/div/div/div[6]/div/table/tbody/tr[4]/td[2]',
'D' : '//*[@id="overview"]/div/div/div/div[5]/div/div/div[1]/p[2]',
'E' : '//*[@id="overview"]/div/div/div/div[5]/div/div/div[4]/p[2]',
'F' : '//*[@id="overview"]/div/div/div/div[5]/div/div/div[3]/p[2]',
'G' : '//*[@id="overview"]/div/div/div/div[5]/div/div/div[2]/p[2]',
'H' : '(//*[local-name() ="tspan"])[4]',
'I' : '(//*[local-name() ="tspan"])[8]',
'J' : '(//*[local-name() ="tspan"])[5]',
'K' : '(//*[local-name() ="tspan"])[7]',
'L' : '//*[@id="keywords"]/div/div/div[2]/div[1]/div/div[2]/div[1]/span[2]',
'M' : '//*[@id="keywords"]/div/div/div[2]/div[1]/div/div[2]/div[2]/span[2]'}
def open_xcel():
    file_name_xcel = u'C:/Users/ACiurba/Desktop/almostUseless/python_web_scraper/similarweb.xlsx'
    wb = load_workbook(file_name_xcel)  # Work Book
    ws = wb["Sheet1"]  # Work Sheet
    s=Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service = s)
    count = 2
    for file in glob.glob("C:/Users/ACiurba/Desktop/similarwebdownloads/*"):
        driver.get(file)
        driver.implicitly_wait(1000)
        for item in x_path_dict:
            location = item + str(count)
            ws[location] = driver.find_element(By.XPATH, x_path_dict[item]).text
        count = count + 1
    wb.save(file_name_xcel)
open_xcel()
