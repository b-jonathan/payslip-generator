import pandas as pd 
from fpdf import FPDF
import os
import mimetypes
import smtplib
from email.message import EmailMessage
class PDF(FPDF):
    #pdf_height = 210
    #pdf_width = 297
    def lines(self):
        self.set_line_width(0.0)
        #self.line(x1,y1,x2,y2)
    # top box
        self.line(20,30,150,30) 
        self.line(20,37,150,37) 
        self.line(20,30,20,37) 
        self.line(150,30,150,37) 

        self.line(20,44,150,44) 
        self.line(20,37,20,44) 
        self.line(150,37,150,44) 

        self.line(20,51,150,51) 
        self.line(20,44,20,51) 
        self.line(150,44,150,51)

    #bottom box
        self.line(20,120,200,120) 
        self.line(20,127,200,127) 
        self.line(20,120,20,127) 
        self.line(200,120,200,127) 
    
    #right box
        self.line(160,13,200,13) 
        self.line(160,44,200,44)    
        self.line(160,23,200,23)
        self.line(160,13,160,44) 
        self.line(180,13,180,44)
        self.line(200,13,200,44) 

    def topBox(self,testdata):
        
        self.set_font('Helvetica','b' ,13)
        self.set_text_color(0,0,0)
        self.set_xy(30,19)
        self.cell(w=20.0, h=5.0, align='L', txt="PT Interflex Sejahtera Perdana", border=0)
        self.set_font('Helvetica','u',11)
        self.set_xy(100,25)
        self.cell(w=20.0, h=5.0, align='L', txt="SLIP GAJI KARYAWAN", border=0)
        self.set_font('Helvetica','' ,11)
        self.set_text_color(0,0,0)
        self.set_xy(20,24)
        self.cell(w=20.0, h=20.0, align='L', txt="GAJI BULAN", border=0)
        self.set_xy(20,31)
        self.cell(w=20.0, h=20.0, align='L', txt="NAMA KARYAWAN & GOL", border=0)
        self.set_xy(20,38)
        self.cell(w=20.0, h=20.0, align='L', txt="DEPT.", border=0)

        self.set_xy(72,24)
        self.cell(w=20.0, h=20.0, align='L', txt= ": SEPTEMBER 2022", border=0)
        self.set_xy(72,31)
        name = testdata["Nama"].upper()
        self.cell(w=20.0, h=20.0, align='L', txt= f": {name}", border=0)
        self.set_xy(72,38)
        bagian = testdata["Bagian"].upper()
        self.cell(w=20.0, h=20.0, align='L', txt= f": {bagian}", border=0)

    def rightBox(self):
        self.set_font('Helvetica','' ,8)
        self.set_text_color(0,0,0)
        self.set_xy(160,10)
        self.cell(w=20.0, h=20.0, align='L', txt= "Dibuat Oleh:", border=0)
        self.set_xy(180,10)
        self.cell(w=20.0, h=20.0, align='L', txt= "Diterima Oleh:", border=0)

    def penghasilan(self,testdata):
        self.set_font('Helvetica','BU' ,11)
        self.set_text_color(0,0,0)
        self.set_xy(20,45)
        self.cell(w=20.0, h=20.0, align='L', txt="PENGHASILAN", border=0)

        self.set_font('Helvetica','' ,11)
        self.set_xy(20,52)
        self.cell(w=20.0, h=20.0, align='L', txt="Gaji Pokok", border=0)
        
        self.set_xy(20,60)
        self.cell(w=20.0, h=20.0, align='L', txt="Tj. Pekerjaan", border=0)
        
        self.set_xy(20,68)
        self.cell(w=20.0, h=20.0, align='L', txt="Tj. Transport", border=0)
        
        self.set_xy(20,76)
        self.cell(w=20.0, h=20.0, align='L', txt="Incentive", border=0)
        
        self.set_xy(20,84)
        self.cell(w=20.0, h=20.0, align='L', txt="Lembur", border=0)
        
        self.set_xy(20,92)
        self.cell(w=20.0, h=20.0, align='L', txt="Shift", border=0)
        
        self.set_xy(30,100)
        self.cell(w=20.0, h=20.0, align='L', txt="Total (A)", border=0)

        for i in range (7):
            n = i*8
            self.set_xy(50,52+n)
            self.cell(w=20.0, h=20.0, align='L', txt="=", border=0)
        gp = testdata["Gaji Pokok"]
        self.set_xy(60,52)
        self.cell(w=20.0, h=20.0, align='L', txt= f"Rp. {gp:,}", border=0)
        tjp = testdata["Tunjangan Pekerjaan"]
        self.set_xy(60,60)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {tjp:,}", border=0)
        tjt = testdata["Tunjangan Transport"]
        self.set_xy(60,68)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {tjt:,}", border=0)

        if testdata["Insentif Kehadiran"] != 0:
            incentive = testdata["Insentif Kehadiran"]
            self.set_xy(60,76)
            self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {incentive:,.0f}", border=0)
        lembur = testdata["Total Lembur"]
        self.set_xy(60,84)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {lembur:,.0f}", border=0)
        ts = testdata["Tunjangan Shift"]
        self.set_xy(60,92)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {ts:,.0f}", border=0)
        tg = testdata["Total Gaji"]
        self.set_xy(60,100)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {tg:,.0f}", border=0)

    def potongan(self,testdata):
        self.set_font('Helvetica','BU' ,11)
        self.set_text_color(0,0,0)
        self.set_xy(100,45)
        self.cell(w=20.0, h=20.0, align='L', txt="POTONGAN", border=0)

        self.set_font('Helvetica','' ,11)
        self.set_xy(100,52)
        self.cell(w=20.0, h=20.0, align='L', txt="PPH 21 & PPH Lembur", border=0)
        
        self.set_xy(100,60)
        self.cell(w=20.0, h=20.0, align='L', txt="Jamsostek & BPJS", border=0)
        
        self.set_xy(100,68)
        self.cell(w=20.0, h=20.0, align='L', txt="Terlambat", border=0)
        
        self.set_xy(100,76)
        self.cell(w=20.0, h=20.0, align='L', txt="Absen", border=0)
        
        self.set_xy(100,84)
        self.cell(w=20.0, h=20.0, align='L', txt="Lainnya", border=0)

        self.set_xy(110,100)
        self.cell(w=20.0, h=20.0, align='L', txt="Total (B)", border=0)

        for i in range (7):
            if i != 5:
                n = i*8
                self.set_xy(150,52+n)
                self.cell(w=20.0, h=20.0, align='L', txt="=", border=0)

        totalPPH = 0
        if pd.notna(testdata["PPH"]):
            totalPPH += testdata["PPH"]
        if pd.notna(testdata["PPH Lembur"]):
            totalPPH += testdata["PPH Lembur"]
        self.set_xy(160,52)
        self.cell(w=20.0, h=20.0, align='L', txt= f"Rp. {totalPPH:,.0f}", border=0)
        
        bpjs = testdata["BPJSTK"]+testdata["BPJSKES"]
        self.set_xy(160,60)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {bpjs:,}", border=0)
        lainnya = testdata["Koperasi"]
        self.set_xy(160,84)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {lainnya:,}", border=0)
        tb = totalPPH+testdata["BPJSTK"]+testdata["BPJSKES"]+testdata["Koperasi"]
        self.set_xy(160,100)
        self.cell(w=20.0, h=20.0, align='L', txt=f"Rp. {tb:,.0f}", border=0)
    
    def bottomBox(self,testdata):
        self.set_font('Helvetica','' ,11)
        self.set_text_color(0,0,0)
        self.set_xy(40,121)
        self.cell(w=20.0, h=5.0, align='L', txt="PENERIMAAN BERSIH (A-B)", border=0)
        self.set_xy(100,121)
        gbt = testdata["Gaji Bersih Transfer"]
        self.cell(w=20.0, h=5.0, align='L', txt=f"Rp. {gbt:,.0f}", border=0)

def createPdf(testdata):
    pdf = PDF(orientation="L", unit="mm",format="A5")
    pdf.add_page()
    pdf.image(r"C:\Users\user\OneDrive\Dokumen\Q_do\Logo.png", 20, 13, 8, 15)
    pdf.lines()
    pdf.topBox(testdata)
    pdf.rightBox()
    pdf.penghasilan(testdata)
    pdf.potongan(testdata)
    pdf.bottomBox(testdata)
    pdf.output(r"C:\Users\user\OneDrive\Dokumen\Q_do\test.pdf",)

def createEmail(name,testdata):
    sender = "payslip@interflex.co.id"
    message = EmailMessage()

    message["From"] = sender
    if pd.notna(testdata["Email"]):
        message["To"] = testdata["Email"]
    else:
        message["To"] = ["ben.jrafflesian1@gmail.com"]
    message["Subject"] = f"Payment Slip {name}"

    body = "Test only"

    message.set_content(body)

    attachment_path = r"C:\Users\user\OneDrive\Dokumen\Q_do\test.pdf"
    attachment_file = os.path.basename(attachment_path)
    mime_type_subtype , _ = mimetypes.guess_type(attachment_path) #guess_type returns a tuple, but only index 0 relevant
    mimetype,subtype = mime_type_subtype.split("/",1) #separate into type and subtype
    with open(attachment_path,"rb") as ap:
        message.add_attachment(ap.read(), maintype=mimetype,subtype=subtype,filename=attachment_file)

    mail_server = smtplib.SMTP_SSL("mail.interflex.co.id") #Make sure to use the correct smtp server
    mail_server.login(sender,"Br@ndon123")
    fail = mail_server.send_message(message)
    mail_server.quit()

    os.remove(attachment_path)

def getData():
    df = pd.read_excel (r"C:\Users\user\OneDrive\Dokumen\Q_do\template_gaji.xlsx") 
    return df

df = getData()
for i in range(2):  
    testdata = df.iloc[i]
    name = testdata["Nama"]
    createPdf(testdata)
    createEmail(name,testdata)

