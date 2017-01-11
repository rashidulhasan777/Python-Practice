#! python3

import requests, bs4, openpyxl, re

stateRegex=re.compile(r',\s\D\D')
zipRegex=re.compile(r'\d\d\d\d\d')
page='https://www.nflpa.com/agents/agentsearch?company=&firstName=&lastName=&city=&state=&zip=&h=true&sort='
row=2

for r in range(2, 82):
    wb=openpyxl.load_workbook('nfl_copy.xlsx')
    sheet=wb.get_sheet_by_name('Sheet1')
    res=requests.get(page)
    res.raise_for_status()
    soup=bs4.BeautifulSoup(res.text, 'html.parser')
    profileAddress=soup.select('tr td a')
    length=len(profileAddress)
    for i in range(0, length, 3):
        pPage='https://www.nflpa.com'+profileAddress[i].get('href')
        try:
            res2=requests.get(pPage)
            res2.raise_for_status()
            soup2=bs4.BeautifulSoup(res2.text, 'html.parser')
        except:
            sheet.cell(row=row, column=1).value=''
            sheet.cell(row=row, column=2).value=''
            sheet.cell(row=row, column=3).value=''
            sheet.cell(row=row, column=4).value=''
            sheet.cell(row=row, column=5).value=''
            sheet.cell(row=row, column=6).value=''
            sheet.cell(row=row, column=7).value=''
            sheet.cell(row=row, column=8).value=''
            sheet.cell(row=row, column=9).value=''
            sheet.cell(row=row, column=10).value=''
            row+=1
            continue
        name=soup2.select('div h1')[0].getText().split()
        firstName=name[0]
        if len(name)>1:
            lastName=name[1]
        else:
            lastName=''
        try:
            company=soup2.select('#Jersey')[0].getText()
            address=soup2.select('#profile-address')[0].getText().replace('\n', ',')[1:]
        except:
            company='Not Found'
            address='Not Found'
        try:
            phone=soup2.select('#WorkPhone a')[0].get('href').replace('tel:', '')
        except:
            phone='Not Found'
        try:
            email=soup2.select('#Email a')[0].get('href').replace('mailto:', '')
        except:
            email='Not Found'
        try:
            cityStateZip=soup2.select('#profile-address div')[len(soup2.select('#profile-address div'))-1].getText()
        except:
            cityStateZip='Not Found'
        city=''
        for x in cityStateZip:
            if x==',':
                break
            else:
                city+=x
        try:
            state=stateRegex.search(cityStateZip).group().replace(', ', '')
            zipCode=zipRegex.search(cityStateZip).group()
        except:
            state='Not Found'
            zipCode='Not Found'
        try:
            currentContract=soup2.select('#ContractNegotiatedCount')[0].getText()
        except:
            currentContract='Not Found'
        sheet.cell(row=row, column=1).value=firstName
        sheet.cell(row=row, column=2).value=lastName
        sheet.cell(row=row, column=3).value=company
        sheet.cell(row=row, column=4).value=address
        sheet.cell(row=row, column=5).value=city
        sheet.cell(row=row, column=6).value=state
        sheet.cell(row=row, column=7).value=zipCode
        sheet.cell(row=row, column=8).value=phone
        sheet.cell(row=row, column=9).value=email
        sheet.cell(row=row, column=10).value=currentContract
        row+=1
    page='https://www.nflpa.com/agents/agentsearch?h=true&page='+str(r)
    wb.save('nfl_copy.xlsx')
    print(r)
    
print('Done') 
    

