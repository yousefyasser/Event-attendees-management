import qrcode, PIL, openpyxl, os.path

wb = openpyxl.load_workbook("task1.xlsx")
ws = wb.active
a = ["name: ", "phone number: ", "email: "]

for i in range(2, ws.max_row+1):
    data = ""

    #reads data from excel
    for j in range(65, 68): 
        cell = str(chr(j))+str(i)
        data += a[j-65] + str(ws[cell].value) + "\n"

    #generates qrcode from data read
    img = qrcode.make(data) 
    
    #writes link to excel sheet
    path = "C:/Users/OS/Desktop/guc coursework/IEEE/"
    title = 'person no. '
    link = path + title + str(i) + ".png"
    img.save(link)
    ws['D'+str(i)] = link 

wb.save('task1.xlsx')
