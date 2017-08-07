#! python3
import  bs4, openpyxl
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui

wb=openpyxl.load_workbook('RC Orders .xlsx')
sheet=wb.get_sheet_by_name('Pricing Section')

#browser=webdriver.Firefox(r'C:\Users\Acer\PycharmProjects\HelloWorld\geckodriver.exe')
browser=webdriver.Firefox()
browser.get('https://www.aliexpress.com/')
r=input()
for row in range(2, 488):
    url=sheet.cell(row=row, column=4).value
    if sheet.cell(row=row-1, column=1).value==sheet.cell(row=row, column=1).value:
        sheet.cell(row=row, column=7).value=sheet.cell(row=row-1, column=7).value
        sheet.cell(row=row, column=8).value=sheet.cell(row=row-1, column=8).value
        continue
    try:
        browser.get(url)
    except:
        sheet.cell(row=row, column=8).value='No URL Found'
        continue
    source = browser.page_source
    soup=bs4.BeautifulSoup(source, 'html.parser')
    try:
        value=soup.select('.p-price')[1].getText()
    except:
        try:
            value=soup.select('.p-price')[0].getText()
        except:
            value='Unable to retrieve the value'
    previousValue=sheet.cell(row=row, column=6).value
    sheet.cell(row=row, column=7).value=value
    
    
    try:
        value=float(value)
        if previousValue==value:
            sheet.cell(row=row, column=8).value='Not Changed'
        elif previousValue>value:
            sheet.cell(row=row, column=8).value='Price Decreased'
        else:
            sheet.cell(row=row, column=8).value='Price Increased'
    except:
        sheet.cell(row=row, column=8).value=''
    wb.save('RC Orders .xlsx')
wb.save('RC Orders .xlsx')
print('Done')
