import re, openpyxl

phoneRegex = re.compile(r'[+]?\d?\d{3}[ -/]*\d{3}\d?[ -/]*\d{3}\d?')

wb1=openpyxl.load_workbook('exampleOfPhoneAndName.xlsx')
sheet1=wb1.get_sheet_by_name('Sheet1')
wb2=openpyxl.load_workbook('first tree row1.xlsx')
sheet2=wb2.get_sheet_by_name('Sheet1')
txtDoc = open('problemRecord.txt', 'w')
for rowAll in range(2, 200):
    phoneString = sheet1.cell(row=rowAll, column=1).value
    nameString = sheet1.cell(row=rowAll, column=2).value

    nameList = nameString.split()
    try:
        firstName = str(nameList[0])
        if(len(nameList[1:]) == 1):
            lastName = nameList[1]
        else:
            lastName = nameList[1]
            for i in nameList[2:]:
                lastName += ' ' + i
    except:
        print("Name spliting error. Row %s" %rowAll)

    sheet2.cell(row=rowAll, column=2).value = firstName
    sheet2.cell(row=rowAll, column=3).value = lastName
    wb2.save('first tree row1.xlsx')

    matchingObject = phoneRegex.search(phoneString)
    if(matchingObject==None):
        try:
            firstIndex = phoneString.index('09')
        except:
            try:
                firstIndex = phoneString.index('+')
            except:
                try:
                    firstIndex = phoneString.index('O9')
                except:
                    txtDoc.write(str(rowAll)+'\n')
                    print(rowAll)

        phoneNumber = phoneString[firstIndex]
        for i in range(firstIndex+1, firstIndex+15):
            try:
                if(phoneString[i].isdigit()):
                    phoneNumber +=phoneString[i]
            except:
                break
    else:
        phoneNumber=''
        for i in matchingObject.group():
            if(i.isdigit()):
                phoneNumber+=i
    sheet2.cell(row=rowAll, column=1).value = phoneNumber
    wb2.save('first tree row1.xlsx')
txtDoc.close()
