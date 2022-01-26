# Import required modules
from lxml import html
import requests
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd
from openpyxl import load_workbook


def get_column_values():
    file_name = u'C:/Users/ACiurba/Desktop/python_web_scraper/testCSV.xlsx'
    wb = load_workbook(file_name)  # Work Book
    ws = wb.get_sheet_by_name('Sheet1')  # Work Sheet
    column = ws['B']  # Column
    
    column_list = [str(column[x].value) for x in range(1,len(column))]
    return column_list

column_list = get_column_values()

driver = webdriver.Firefox(executable_path=r'C:/Users/ACiurba/Downloads/geckodriver-v0.30.0-win64/geckodriver.exe')
driver.get("https://www.similarweb.com/")
for item in column_list:
    inputElement = driver.find_element_by_id("search")
    inputElement.send_keys(item)
    inputElement.send_keys(Keys.ENTER)
    driver.navigate().back();
