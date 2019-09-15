# -*- coding: utf-8 -*-
"""
Created on Sat Sep 14 21:03:27 2019

@author: olp16101
"""

import requests
import xlsxwriter
from lxml import etree



def get_rates(year):
    """ The function takes a four-digit numeric year (1990 - 2019) as an
    argument and retrievs the daily Treasury yield curve rates for the
    specified year from the US Treasury website in XML format. The function 
    then return a 2-dimensional list with rates where each column is a yield
    rate type and each row is a date."""

    if year not in range(1990,2020):
        print("Error: no data available for this year")
        return []
    
    url = "https://data.treasury.gov/feed.svc/DailyTreasuryYieldCurveRate" \
    + "Data?$filter=year(NEW_DATE)%20eq%20" + str(year)

    xml_string = requests.get(url).content
    
    tree = etree.fromstring(xml_string)
    
    tbill_date = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                   + "/08/dataservices}NEW_DATE")
    tbill_1month = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_1MONTH")
    tbill_2month = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_2MONTH")
    tbill_3month = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_3MONTH")
    tbill_6month = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_6MONTH")
    tbill_1year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                    + "/08/dataservices}BC_1YEAR")
    tbill_2year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                    + "/08/dataservices}BC_2YEAR")
    tbill_3year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                    + "/08/dataservices}BC_3YEAR")
    tbill_5year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                    + "/08/dataservices}BC_5YEAR")
    tbill_7year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                    + "/08/dataservices}BC_7YEAR")
    tbill_10year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_10YEAR")
    tbill_20year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_20YEAR")
    tbill_30year = tree.findall(".//{http://schemas.microsoft.com/ado/2007" \
                                     + "/08/dataservices}BC_30YEAR")
    
    rates = []
    
    for n in range(len(tbill_date)):
        row = [tbill_date[n].text, tbill_1month[n].text, tbill_2month[n].text, \
              tbill_3month[n].text, tbill_6month[n].text, tbill_1year[n].text, \
              tbill_2year[n].text, tbill_3year[n].text, tbill_5year[n].text, \
              tbill_7year[n].text, tbill_10year[n].text, tbill_20year[n].text, \
              tbill_30year[n].text]
        
        rates.append(row)
        
    return rates









start_year = 2015
end_year = 2017

rates = []

for year in range(start_year, end_year + 1):
    rates += get_rates(year)




file_name = "YieldCurveRates%s-%s.xlsx" % (start_year, end_year)
workbook = xlsxwriter.Workbook(file_name, {'strings_to_numbers':  True})
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Date')
worksheet.write('B1', '1 Mo')
worksheet.write('C1', '2 Mo')
worksheet.write('D1', '3 Mo')
worksheet.write('E1', '6 Mo')
worksheet.write('F1', '1 Yr')
worksheet.write('G1', '2 Yr')
worksheet.write('H1', '3 Yr')
worksheet.write('I1', '5 Yr')
worksheet.write('J1', '7 Yr')
worksheet.write('K1', '10 Yr')
worksheet.write('L1', '20 Yr')
worksheet.write('M1', '30 Yr')

row = 1
col = 0

for date, mo1, mo2, mo3, mo6, ye1, ye2, ye3, ye5, ye7, ye10, ye20, ye30 in \
rates:
    worksheet.write(row, col, date)
    worksheet.write(row, col + 1, mo1)
    worksheet.write(row, col + 2, mo2)
    worksheet.write(row, col + 3, mo3)
    worksheet.write(row, col + 4, mo6)
    worksheet.write(row, col + 5, ye1)
    worksheet.write(row, col + 6, ye2)
    worksheet.write(row, col + 7, ye3)
    worksheet.write(row, col + 8, ye5)
    worksheet.write(row, col + 9, ye7)
    worksheet.write(row, col + 10, ye10)
    worksheet.write(row, col + 11, ye20)
    worksheet.write(row, col + 12, ye30)
    
    row += 1
    
workbook.close()

