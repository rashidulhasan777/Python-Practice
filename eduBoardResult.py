#!python3

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import selenium.webdriver.support.ui as ui
import re, time, bs4, smtplib
captchaRegex=re.compile(r'([0-9] \+ [0-9])')

rollNumber = #Enter roll number
regNumber = #Enter reg number
yearExam = 2019
boardExam= "Rajshahi"




browser=webdriver.Chrome()
browser.get('http://www.educationboardresults.gov.bd/')
while True:
    '''exam = browser.find_element_by_id('exam')
    exam.send_keys('SSC/Dakhil')'''
    year=browser.find_element_by_id('year')
    year.send_keys(yearExam)
    board=browser.find_element_by_id('board')
    board.send_keys(boardExam)
    roll=browser.find_element_by_id('roll')
    roll.send_keys(rollNumber)
    reg=browser.find_element_by_id('reg')
    reg.send_keys(regNumber)
    source=browser.page_source
    captcha=captchaRegex.findall(source)
    sum = int(captcha[0][0]) + int(captcha[0][-1])
    value_s=browser.find_element_by_id('value_s')
    value_s.send_keys(sum)
    submitButton = browser.find_element_by_id('button2')
    submitButton.click()
    time.sleep(5)
    source=browser.page_source
    try:
        soup=bs4.BeautifulSoup(source, 'html.parser')
        result=soup.select('td.black12bold')[1].getText()
        smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login('yourmail@gmail.com', 'password')
        smtpObj.sendmail('yourmail@gmail.com', 'receivermail@gmail.com', '''Subject: Result.\n Result for ''' + str(rollNumber) + ' : is GPA : ' +str(result))
        smtpObj.quit()
        break
    except:
        continue
