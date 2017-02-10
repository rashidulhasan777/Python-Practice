#! python3
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
import bs4, openpyxl, time


url='http://www.trucking.org/'
row=3
wb=openpyxl.load_workbook('2017 TMC Carriers Membership List..xlsx')
sheet=wb.get_sheet_by_name('Sheet1')
browser=webdriver.Chrome()
browser.get(url)
linkElem=browser.find_element_by_link_text('SIGN IN')
linkElem.click()
emailElem=browser.find_element_by_id('LoginTextBox')
emailElem.send_keys('msteele@gbmsummits.com')
passwordElem=browser.find_element_by_id('PasswordTextBox')
passwordElem.send_keys('steele1234')
passwordElem.send_keys(Keys.ENTER)

directoriesElem=browser.find_element_by_link_text('DIRECTORIES')
directoriesElem.click()
tmsElem=browser.find_element_by_link_text('TMC Membership Directory')
tmsElem.click()

#browser.get('http://www.trucking.org/CouncilDirectory.aspx?council=TMC')
browser.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder_ddlLevel1"]').send_keys('Fleet')
browser.find_element_by_id('ctl00_ContentPlaceHolder_btnSearch').click()
for m in range(42):
    source = browser.page_source
    soup=bs4.BeautifulSoup(source, 'html.parser')
    links=soup.select('.inner div a')

    for i in range(10, 30):
        browser.get(url+links[i].get('href'))
        mainSource=browser.page_source
        soup2=bs4.BeautifulSoup(mainSource, 'html.parser')
        sheet.cell(row=row, column=1).value=soup2.select('#ctl00_ContentPlaceHolder_lblCompany')[0].getText()
        sheet.cell(row=row, column=2).value=soup2.select('#ctl00_ContentPlaceHolder_lblTitle')[0].getText()
        sheet.cell(row=row, column=3).value=soup2.select('#ctl00_ContentPlaceHolder_lblLevel1')[0].getText()
        sheet.cell(row=row, column=4).value=soup2.select('#ctl00_ContentPlaceHolder_lblLevel2')[0].getText()
        sheet.cell(row=row, column=5).value=soup2.select('#ctl00_ContentPlaceHolder_lblOfficePhone')[0].getText()
        sheet.cell(row=row, column=6).value=soup2.select('#ctl00_ContentPlaceHolder_lblEmail')[0].getText()
        row+=1
    for i in range(20):
        browser.back()
    x=browser.find_element_by_id('ctl00_ContentPlaceHolder_dpListView_ctl00_NextButton')
    x.click()
    if m>23 and m<27:
        time.sleep(30)
    
    wb.save('2017 TMC Carriers Membership List..xlsx')

print('Done')

