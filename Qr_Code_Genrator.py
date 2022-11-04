from csv import writer
from doctest import Example
from genericpath import isfile
import pandas as pd
import qrcode
import os.path
from PIL import Image
import stat

# Consts

File_Path=os.path.abspath('./example.xlsx')
Images_Path="./Images/"
Image_Title='Person no. '
Logo_Path=os.path.abspath("./Images/logo.png")

class QR_Genrator:
    def __init__(self , path,logo,save,title) :
        exist=isfile(path)
        if not exist:
            print("[ERORR]: File does not exist. Path entered is : "+ path)
            exit()

        self.logo=logo    
        self.path=path
        self.save=save
        self.title=title
        self.links_list=[]
        self.read()
        self.Qr_genrator()
        self.write()


    def read(self):
        self.data=pd.read_excel(self.path)
        # print(self.data.loc[0])
        self.no_of_rows=len(self.data) 
        self.no_of_columns=len(list(self.data))
    

    def get_data(self,i):
        data_list=str(self.data.loc[i]).splitlines()
        # print(data_list)
        data=''
        for element in range(len(data_list)-1):
            data+=(data_list[element]+'\n') 
        return data    

    def Genrate_img(self,qr):
        logo = Image.open(self.logo)
        basewidth = 100
        wpercent = (basewidth/float(logo.size[0]))
        hsize = int((float(logo.size[1])*float(wpercent)))
        logo = logo.resize((basewidth, hsize), Image.ANTIALIAS)
        img=qr.make_image(fill_color='black', back_color="white").convert('RGB')
        pos = ((img.size[0] - logo.size[0]) // 2,(img.size[1] - logo.size[1]) // 2)
        img.paste(logo, pos)
 
        return img

    def Qr_genrator(self):
        for i in range(self.no_of_rows):
            
            QRcode = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H)
            
            data=self.get_data(i)
            print(data)
            QRcode.add_data(data)

            img=self.Genrate_img(QRcode)

            link=self.save+self.title+str(i)+".png"
            
            self.links_list.append(link)
            
            img.save(link)
        return self.links_list    
   
    def write(self):
        path=os.path.abspath(self.path)
        col_name=["QR_Link"]
        data=self.links_list
        df=pd.DataFrame(data,columns=col_name)
        os.chmod(path,stat.S_IRWXO)
        writer=pd.ExcelWriter(path,engine='xlsxwriter')
        df.to_excel(writer,sheet_name='Sheet1')
        writer.save()
       



def main ():
    Genrator =QR_Genrator(File_Path,Logo_Path,Images_Path,Image_Title)
if __name__=="__main__":
    main()
    input()

