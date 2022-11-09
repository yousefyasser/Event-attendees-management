import qrcode, openpyxl, os
from faker import Faker
from faker.providers import DynamicProvider
from PIL import Image
from sys import argv

wb = openpyxl.Workbook()
ws = wb.active
headers = ['Name', 'Phone number', 'Email', 'Attendee Group', 'QR link', 'ID']


def create_ExcelSheet():
    fake = Faker()
        
    groups_provider = DynamicProvider(
         provider_name = "group_provider",
         elements=["L", "LS", "LD"],
    )
    fake.add_provider(groups_provider)
    
    for col in range(65, 71):
        ws[chr(col)+'1'] = headers[col-65]

    try:
        AttendanceNumber = int(input("Enter number of attendees: "))
        for i in range(2, AttendanceNumber+2):
            ws['A'+str(i)] = fake.name()
            ws['B'+str(i)] = fake.phone_number()
            ws['C'+str(i)] = fake.email()
            ws['D'+str(i)] = fake.group_provider()
    except:
        print("Please enter a number")


def Generate_img(qr):
        logo = Image.open('logo.png')
        basewidth = 100
        wpercent = (basewidth/float(logo.size[0]))
        hsize = int((float(logo.size[1])*float(wpercent)))
        logo = logo.resize((basewidth, hsize), Image.ANTIALIAS)
        img=qr.make_image(fill_color='black', back_color="white").convert('RGB')
        pos = ((img.size[0] - logo.size[0]) // 2,(img.size[1] - logo.size[1]) // 2)
        img.paste(logo, pos)
 
        return img
    

def main(imagesPath):
    os.makedirs(imagesPath, exist_ok = True)
    
    create_ExcelSheet()
    for i in range(2, ws.max_row+1):
        data = ""

        #reads data from excel
        for j in range(65, 69): 
            cell = chr(j)+str(i)
            data += headers[j-65] + ': ' + str(ws[cell].value) + "\n"

        #generates qrcode from data read
        QRcode = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
        QRcode.add_data(data)
        img = Generate_img(QRcode)
        
        #modifies the excel sheet
        title = 'person no. '
        link = os.path.join(imagesPath, f"{title}{str(i-1)}.png")
        img.save(link)
        ws['E'+str(i)] = link
        ws['F'+str(i)] = 'IEEE-'+str(i-1);


if __name__ == "__main__":
    imagesPath = os.path.dirname(argv[0]) if len(os.path.dirname(argv[0])) != 0 else os.getcwd()    
    imagesPath = os.path.join(imagesPath, 'images')
    main(imagesPath)
    wb.save('Sample.xlsx')
