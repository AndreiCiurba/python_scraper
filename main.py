# Import required modules
from lxml import html, etree
import requests
import email
# import pandas as pd
from openpyxl import load_workbook

from bs4 import BeautifulSoup
from html_table_extractor.extractor import Extractor


    #
    # #setup web
    # #add loop through all files
    #
    # #for file in all files:
    file_name = u'C:/Users/ACiurba/Desktop/similarwebdownloads/test.html'
    count = 2
    html_report_part1 = open(file_name,'r', errors="ignore")
    soup = BeautifulSoup(html_report_part1, 'html.parser')
    html_string = soup.prettify()
    start = 'wa-overview__engagement-value">'
    end = '</p>'

    print((html_string.split(start))[1].split(end)[0])
    total_visits_location = 'F' + str(count)
    ws[total_visits_location] = "asda"

                # for table in soup.body.find_all("p", recursive=False):
                #     extractor = Extractor(table)
                #     extractor.parse()
                #     print(extractor.return_list())


    #
    #
    # distributor = soup.find("p", {"class": '3D"wa-overview__title"'}).get_text().split('<=')[0]
    # print(soup.find("tspan", {"class": '3D"wa-traffic-sources__channels-data-label"'}).get_text())

    #
    #
    # url = soup.find("p", {"class": '3D"wa-overview__title"'}).get_text().split('<=')[0]
    # print(soup.find("tspan", {"class": '3D"wa-traffic-sources__channels-data-label"'}).get_text())

    #
    # country = soup.find("p", {"class": '3D"wa-overview__title"'}).get_text().split('<=')[0]
    # print(soup.find("tspan", {"class": '3D"wa-traffic-sources__channels-data-label"'}).get_text())

    #
    #
    #
    # total_visits = soup.findAll("p", {"class": '3D"wa-overview__engagement-value"'})[0].get_text().split('<=')[0]
    # print(soup.findAll("p", {"class": '3D"wa-overview__engagement-value"'})[0].get_text().split('<=')[0]




    #
    # wb.save(file_name_xcel)
open_xcel()



#
# all fara search din tabel
# alea 4 de sus
# Organic vs. Paid
