#! python3
import requests, bs4, openpyxl

mainUrl='https://www.kickstarter.com'
url='https://www.kickstarter.com/discover/advanced?woe_id=23424977&goal=2&raised=0&sort=newest&seed=2477429&page='
wb=openpyxl.load_workbook('Kscrape.xlsx')
sheet=wb.get_sheet_by_name('Sheet1')
row=2
for i in range(1, 201):
    try:
        res=requests.get(url+str(i))
        soup=bs4.BeautifulSoup(res.text, 'html.parser')
        links=soup.select('.project-thumbnail-wrap')
    except:
        row+=20
        continue
    length=len(links)
    for j in range(length):
        try:
            res2=requests.get(mainUrl+links[j].get('href'))
            soup2=bs4.BeautifulSoup(res2.text, 'html.parser')
            res3=requests.get(mainUrl+soup2.select('.remote_modal_dialog')[0].get('href'))
            soup3=bs4.BeautifulSoup(res3.text, 'html.parser')
        except:
            row+=1
            continue
        try:
            sheet.cell(row=row, column=1).value=soup2.select('h2')[0].getText().replace('\n', '') # A
        except:
            sheet.cell(row=row, column=1).value=''
        try:
            sheet.cell(row=row, column=2).value=mainUrl+links[j].get('href') #B
        except:
            sheet.cell(row=row, column=2).value=''
        try:
            sheet.cell(row=row, column=3).value=soup2.select('.money')[1].getText()
        except:
            sheet.cell(row=row, column=3).value=''
        try:
            sheet.cell(row=row, column=4).value=soup2.select('#pledged + span')[0].getText()
        except:
            sheet.cell(row=row, column=4).value=''
        try:
            sheet.cell(row=row, column=5).value=soup2.select('a[class="nowrap navy-700 medium type-12"]')[0].getText().replace('\n', '')[:len(soup2.select('a[class="nowrap navy-700 medium type-12"]')[0].getText().replace('\n', ''))-4]
        except:
            sheet.cell(row=row, column=5).value=''
        try:
            sheet.cell(row=row, column=6).value=soup2.select('a[class="nowrap navy-700 medium type-12"]')[0].getText().replace('\n', '')[len(soup2.select('a[class="nowrap navy-700 medium type-12"]')[0].getText().replace('\n', ''))-2:]
        except:
            sheet.cell(row=row, column=6).value=''
        try:
            sheet.cell(row=row, column=7).value=soup2.select('div[class="js-backers_count block type-16 type-24-md medium navy-700"]')[0].getText().replace('\n', '')
        except:
            sheet.cell(row=row, column=7).value=''
        try:
            sheet.cell(row=row, column=8).value=str(int((int(soup2.select('#project_duration_data')[0]['data-hours-remaining'])/24)))
        except:
            sheet.cell(row=row, column=8).value=''
        try:
            counter=0
            tempCatagory=soup2.select('div[class="flex items-center scroll-x"]')[0].getText()
            for k in range(len(tempCatagory)):
                if tempCatagory[k]=='\n':
                    counter+=1
                if counter==5:
                    break
            catagory=tempCatagory[:k].replace('\n', '')
            sheet.cell(row=row, column=9).value=catagory
        except:
            sheet.cell(row=row, column=9).value=''
        try:
            sheet.cell(row=row, column=10).value=mainUrl+soup2.select('.remote_modal_dialog')[0].get('href')
        except:
            sheet.cell(row=row, column=10).value=''
        try:
            sheet.cell(row=row, column=11).value=soup3.select('.green-dark')[1].getText().replace('\n', '').split()[0]
        except:
            sheet.cell(row=row, column=11).value=''
        try:
            sheet.cell(row=row, column=12).value=soup3.select('.green-dark')[1].getText().replace('\n', '').split()[1]
        except:
            sheet.cell(row=row, column=12).value=''
        try:
            counter=0
            createdBaked=soup3.select('div[class="created-projects py2 f5 mb3"]')[0].getText()
            for l in range(len(createdBaked)):
                if createdBaked[l]=='\n':
                    counter+=1
                if counter==3:
                    break
            created=createdBaked[:l].replace('\n', '')
            baked=createdBaked[l+3:].replace('\n', '')
            sheet.cell(row=row, column=13).value=created
            sheet.cell(row=row, column=14).value=baked
        except:
            sheet.cell(row=row, column=13).value=''
            sheet.cell(row=row, column=14).value=''
        try:
            sheet.cell(row=row, column=15).value=soup3.select('div[class="last-login py2 border-bottom f5"] > time')[0].getText()
        except:
            sheet.cell(row=row, column=15).value=''
        try:
            sheet.cell(row=row, column=16).value=soup3.select('div[class="facebook py2 border-bottom f5"]')[0].getText().replace('\n', '')
        except:
            sheet.cell(row=row, column=16).value=''
        try:
            sheet.cell(row=row, column=17).value=soup3.select('div[class="flag-body"]')[0].getText().split()[0]
        except:
            sheet.cell(row=row, column=17).value=''
        try:
            sheet.cell(row=row, column=18).value=soup3.select('div[class="flag-body"]')[0].getText().split()[1]
        except:
            sheet.cell(row=row, column=18).value=''
        try:
            sheet.cell(row=row, column=19).value=mainUrl+soup3.select('div[class="flag-body"] > a')[0].get('href')
        except:
            sheet.cell(row=row, column=19).value=''
        try:
            sheet.cell(row=row, column=20).value=soup3.select('div[class="flag-body"]')[1].getText().split()[0]
        except:
            sheet.cell(row=row, column=20).value=''
        try:
            sheet.cell(row=row, column=21).value=soup3.select('div[class="flag-body"]')[1].getText().split()[1]
        except:
            sheet.cell(row=row, column=21).value=''
        try:
            sheet.cell(row=row, column=22).value=mainUrl+soup3.select('div[class="flag-body"] > a')[1].get('href')
        except:
            sheet.cell(row=row, column=22).value=''
        try:
            sheet.cell(row=row, column=23).value=soup3.select('div[class="flag-body"]')[2].getText().split()[0]
        except:
            sheet.cell(row=row, column=23).value=''
        try:
            sheet.cell(row=row, column=24).value=soup3.select('div[class="flag-body"]')[2].getText().split()[1]
        except:
            sheet.cell(row=row, column=24).value=''
        try:
            sheet.cell(row=row, column=25).value=mainUrl+soup3.select('div[class="flag-body"] > a')[2].get('href')
        except:
            sheet.cell(row=row, column=25).value=''
        try:
            sheet.cell(row=row, column=26).value=soup3.select('div[class="flag-body"]')[3].getText().split()[0]
        except:
            sheet.cell(row=row, column=26).value=''
        try:
            sheet.cell(row=row, column=27).value=soup3.select('div[class="flag-body"]')[3].getText().split()[1]
        except:
            sheet.cell(row=row, column=27).value=''
        try:
            sheet.cell(row=row, column=28).value=mainUrl+soup3.select('div[class="flag-body"] > a')[3].get('href')
        except:
            sheet.cell(row=row, column=28).value=''
        try:
            sheet.cell(row=row, column=29).value=soup3.select('ul[class="links list f5 bold"] > li > a')[0].get('href')
        except:
            sheet.cell(row=row, column=29).value=''
        try:
            sheet.cell(row=row, column=32).value=soup3.select('ul[class="links list f5 bold"] > li > a')[1].get('href')
        except:
            sheet.cell(row=row, column=32).value=''
        try:
            sheet.cell(row=row, column=35).value=soup3.select('ul[class="links list f5 bold"] > li > a')[2].get('href')
        except:
            sheet.cell(row=row, column=35).value=''
        try:
            sheet.cell(row=row, column=38).value=soup3.select('ul[class="links list f5 bold"] > li > a')[3].get('href')
        except:
            sheet.cell(row=row, column=38).value=''
        row+=1
    wb.save('Kscrape.xlsx')
    print(row)
print('Done')    
