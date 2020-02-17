import bitly_api, openpyxl, time

wb = openpyxl.load_workbook('fileName')
sheet = wb['Youtube']
BITLY_ACCESS_TOKEN ="" 
b = bitly_api.Connection(access_token = BITLY_ACCESS_TOKEN)
for i in range(2, 97):
    longurl = sheet.cell(row=i, column=3).value.strip()
    response = b.shorten(longurl) 
    sheet.cell(row=i, column=4).value=response['url']
    time.sleep(1)
wb.save('fileName')
