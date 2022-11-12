import qrcode, openpyxl, os, random, json, warnings, sys
from faker import Faker
from PIL import Image
from genericpath import isfile
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

START_COL = ord('A')
END_COL = ord('G')
START_ROW = 2

# create a folder "images" and store its path
workingDir = sys.path[0]
os.chdir(workingDir)
imgPath = os.path.join(workingDir, 'images')
os.makedirs(imgPath, exist_ok=True)

# connect to google drive api
googleAuth = GoogleAuth()
googleAuth.LocalWebserverAuth()


class QrExcel:
    def __init__(self, imgPath):
        self.headers = ['Name', 'Phone number', 'Email', 'Attendee Group', 'QR link', 'ID']
        self.imgPath = imgPath
        self.drive = GoogleDrive(googleAuth)
        self.main()

    def upload_to_drive(self, row: int, path: str) -> str:
        file = self.drive.CreateFile()
        file.SetContentFile(path)
        file['title'] = f'person no. {row - 1}'
        file.Upload()

        return file['alternateLink']

    def read_existing_file(self, file):
        workbook = openpyxl.open(file)
        worksheet = workbook.active

        for row in range(START_ROW, worksheet.max_row + 1):
            img, path = self.create_qr(row, worksheet)
            link = self.upload_to_drive(row, path)
            worksheet[f'E{row}'] = link
            worksheet[f'F{row}'] = f'IEEE-{str(row - 1)}'

        workbook.save(file)

    def dummy_data(self):
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

        fake = Faker()
        elements = ['L', 'LS', 'LD']

        for col in range(START_COL, END_COL):
            worksheet[f'{chr(col)}1'] = self.headers[col - START_COL]

        attendanceNum = input('enter number of attendees: ')
        if attendanceNum.isdigit():
            attendanceNum = int(attendanceNum)
            if attendanceNum < 0 or attendanceNum > 10:
                print('please enter a number between 0 and 10')
                quit()
        else:
            print('please enter a number')
            quit()

        attendanceNum = int(attendanceNum)
        for row in range(START_ROW, attendanceNum + START_ROW):
            worksheet[f'A{row}'] = fake.name()
            worksheet[f'B{row}'] = fake.phone_number()
            worksheet[f'C{row}'] = fake.email()
            worksheet[f'D{row}'] = random.choice(elements)

            img, path = self.create_qr(row, worksheet)
            link = self.upload_to_drive(row, path)

            worksheet[f'E{row}'] = link
            worksheet[f'F{row}'] = f'IEEE-{str(row - 1)}'

        workbook.save('Sample.xlsx')

    def create_qr(self, row: int, ws) -> (Image.Image, str):
        # reads data from excel
        data = ""

        for j in range(START_COL, END_COL - 2):
            cell = chr(j) + str(row)
            data += f'{self.headers[j - START_COL]}: {str(ws[cell].value)}\n'

        # generates qrcode from data read
        QRcode = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
        json_data = json.loads(json.dumps(data))
        QRcode.add_data(json_data)
        img = self.place_logo(QRcode)

        # save qrcode in images folder
        title = 'person no. '
        link = os.path.join(imgPath, f'{title}{str(row - 1)}.png')
        img.save(link)

        return img, link

    def place_logo(self, qr: qrcode.QRCode) -> Image.Image:
        logo = Image.open('logo.png')
        baseWidth = 100
        widthPercent = (baseWidth / float(logo.size[0]))
        heightSize = int((float(logo.size[1]) * float(widthPercent)))
        logo = logo.resize((baseWidth, heightSize), Image.ANTIALIAS)
        img = qr.make_image(fill_color='black', back_color='white').convert('RGB')
        pos = ((img.size[0] - logo.size[0]) // 2, (img.size[1] - logo.size[1]) // 2)
        img.paste(logo, pos)

        return img

    def main(self):
        choice = input('enter excel file name (without extension) or \'any\' to generate dummy data: ')
        self.dummy_data() if choice == 'any' else self.read_existing_file(f'{choice}.xlsx') if isfile(f'{choice}.xlsx') else print('invalid input')


if __name__ == '__main__':
    warnings.simplefilter('ignore')
    QrExcel(imgPath)
