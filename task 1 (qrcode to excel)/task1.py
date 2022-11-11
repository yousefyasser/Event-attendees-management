import qrcode, openpyxl, os, random, json
from faker import Faker
from PIL import Image
from sys import argv
from genericpath import isfile
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

START_COL = ord('A')
END_COL = ord('G')
START_ROW = 2
headers = ['Name', 'Phone number', 'Email', 'Attendee Group', 'QR link', 'ID']

# connect to google drive api
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

# create a folder "images" and store its path
imgPath = os.path.dirname(argv[0]) if len(os.path.dirname(argv[0])) != 0 else os.getcwd()
imgPath = os.path.join(imgPath, 'images')
os.makedirs(imgPath, exist_ok=True)


def upload_to_drive(path: str) -> str:
    file = drive.CreateFile()
    file.SetContentFile(path)
    file.Upload()

    return file['alternateLink']


def read_existing_file(file):
    workbook = openpyxl.open(file)
    worksheet = workbook.active

    for row in range(START_ROW, worksheet.max_row + 1):
        img, path = create_qr(row, worksheet)
        link = upload_to_drive(path)
        worksheet[f'E{row}'] = link
        worksheet[f'F{row}'] = f'IEEE-{str(row - 1)}'

    workbook.save(file)


def dummy_data():
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    fake = Faker()
    elements = ["L", "LS", "LD"]

    for col in range(START_COL, END_COL):
        worksheet[f'{chr(col)}1'] = headers[col - START_COL]

    try:
        AttendanceNumber = int(input("Enter number of attendees: "))

        for i in range(START_ROW, AttendanceNumber + START_ROW):
            worksheet[f'A{i}'] = fake.name()
            worksheet[f'B{i}'] = fake.phone_number()
            worksheet[f'C{i}'] = fake.email()
            worksheet[f'D{i}'] = random.choice(elements)
            img, path = create_qr(i, worksheet)
            link = upload_to_drive(path)
            worksheet[f'E{i}'] = link
            worksheet[f'F{i}'] = f'IEEE-{str(i - 1)}'
    except:
        print("Please enter a number")

    workbook.save('Sample.xlsx')


def create_qr(row: int, ws) -> (Image, str):
    # reads data from excel
    data = ""

    for j in range(START_COL, END_COL - 2):
        cell = chr(j) + str(row)
        data += f'{headers[j - START_COL]}: {str(ws[cell].value)}\n'

    # generates qrcode from data read
    QRcode = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
    json_data = json.loads(json.dumps(data))
    QRcode.add_data(json_data)
    img = place_logo(QRcode)

    # save qrcode in images folder
    title = 'person no. '
    link = os.path.join(imgPath, f'{title}{str(row - 1)}.png')
    img.save(link)

    return img, link


def place_logo(qr: Image) -> Image:
    logo = Image.open('logo.png')
    baseWidth = 100
    widthPercent = (baseWidth / float(logo.size[0]))
    heightSize = int((float(logo.size[1]) * float(widthPercent)))
    logo = logo.resize((baseWidth, heightSize), Image.ANTIALIAS)
    img = qr.make_image(fill_color='black', back_color="white").convert('RGB')
    pos = ((img.size[0] - logo.size[0]) // 2, (img.size[1] - logo.size[1]) // 2)
    img.paste(logo, pos)

    return img


def main():
    choice = input("enter excel file name (without extension) or \'any\' to generate dummy data: ")
    dummy_data() if choice == 'any' else read_existing_file(f'{choice}.xlsx') if isfile(f'{choice}.xlsx') else print('invalid input')


if __name__ == "__main__":
    main()
