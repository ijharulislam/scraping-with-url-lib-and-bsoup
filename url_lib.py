import urllib2

from lxml import html
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re
from datetime import datetime
import json


response = urllib2.urlopen('https://www.radioshack.com/apps/store-locator/get_surrounding_stores.php?shop=radioshack-demo.myshopify.com&latitude=0&longitude=0&max_distance=0&limit=100&calc_distance=1')
json_data = json.load(response) 
data = []
first_output = [] 
store_list = json_data["stores"]
for i in store_list:
	output = {}
	output["Store Number"] = i["store_id"]
	soup = BeautifulSoup(i["4"], "lxml")
	output["Store Name"] = soup.find("body").find("span", class_="name").text
	output["City"] = soup.find("body").find("span", class_="city").text
	output["Street Adress"] = soup.find("body").find("span", class_="address").text
	output["State"] = soup.find("body").find("span", class_="prov_state").text
	output["Zip Code"] = soup.find("body").find("span", class_="postal_zip").text

	data.append(output)
print data 





def write_to_excel(workbook,worksheet,data):
        
        # w = tzwhere.tzwhere()
        bold = workbook.add_format({'bold': True})
        bold_italic = workbook.add_format({'bold': True, 'italic':True})
        border_bold = workbook.add_format({'border':True,'bold':True})
        border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
        border = workbook.add_format({'border':True,'bold':True})
        
        #worksheet = workbook.add_worksheet('%s_%s'%(a,j))
        worksheet.set_column('B:D', 22)
        worksheet.set_column('E:F', 33)
        row = 0
        col = 0


        worksheet.write(row,col,'Store List',bold)
        row = row + 1

        row = row + 2

        worksheet.write(row,col,'Sl No',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Store No.',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Store Name',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Street Address',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'City',border_bold_grey)
    
        col = col + 1
        worksheet.write(row,col,'State',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Zip Code',border_bold_grey)
    

        row = row + 1
        i = 0


        """{'City': u' HOMER', 'Store Name': u' TECH CONNECT, INC', 
        'Zip Code': u' 99603', 'Street Adress': u' 432 EAST PIONEER AVE #C', 
        'State': u' AK', 'Store Number': u'2766977'}"""

        for output in data:
                
            i = i + 1
            col = 0
            worksheet.write(row, col, i, border)
            col = col + 1
            worksheet.write(row, col, output["Store Number"] if output.has_key('Store Number') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Store Name"] if output.has_key('Store Name') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Street Adress"] if output.has_key('Street Adress') else '',border)
            col = col + 1
            worksheet.write(row, col, output["City"] if output.has_key('City') else '',border)
            col = col + 1
            worksheet.write(row, col, output["State"] if output.has_key('State') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Zip Code"] if output.has_key('Zip Code') else '',border)

            col = col + 1
            row = row + 1

workbook = xlsxwriter.Workbook('store-list.xlsx')
worksheet = workbook.add_worksheet('Store List')
write_to_excel(workbook,worksheet,data)
workbook.close()
