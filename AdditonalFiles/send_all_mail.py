import xlrd
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from email import encoders
from email.mime.base import MIMEBase

import AdditonalFiles.banner as bb
import AdditonalFiles.sales_trend_table as st
import AdditonalFiles.seen_rx_table as sr
import AdditonalFiles.doctor_call_table as dc


def send_all_report():
    bb.all_banner()
    import AdditonalFiles.all_kpi_image as all_kpi
    import AdditonalFiles.separate_all_total_data as total
    # total.seperate_total_data()
    all_kpi.create_total_kpi()
    all_kpi.all_kpi_images()

    # ------------ Group email ----------------------------------------
    msgRoot = MIMEMultipart('related')
    me = 'erp-bi.service@transcombd.com'

    to = ['', 'rejaul.islam@transcombd.com']
    cc = ['', '']
    bcc = ['', '']

    # to = ['', 'tafsir.bashar@skf.transcombd.com']
    # cc = ['', '']
    # bcc = ['yakub@transcombd.com', 'rejaul.islam@transcombd.com']

    recipient = to + cc + bcc
    # print('mail Sending to = ', email)

    subject = "Brand Performance Report "

    email_server_host = 'mail.transcombd.com'
    port = 25

    msgRoot['From'] = me

    msgRoot['To'] = ', '.join(to)
    msgRoot['Cc'] = ', '.join(cc)
    msgRoot['Bcc'] = ', '.join(bcc)
    msgRoot['Subject'] = subject

    msg = MIMEMultipart()
    msgRoot.attach(msg)

    ## Banner Image
    b = open("./Images/Banner/all_banner_ai.png", 'rb')
    banner = MIMEImage(b.read())
    b.close()
    banner.add_header('Content-ID', '<banner>')
    msgRoot.attach(banner)

    # # All KPI
    total_kpi = open("./Images/total_kpi_images.png", 'rb')
    total_kpi = MIMEImage(total_kpi.read())
    total_kpi.add_header('Content-ID', '<total_kpi>')
    msgRoot.attach(total_kpi)

    # # KPI Image
    img = open("./Images/all_kpi_image.png", 'rb')
    img1 = MIMEImage(img.read())
    img.close()
    img1.add_header('Content-ID', '<img1>')
    msgRoot.attach(img1)

    msgText = MIMEText("""

    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <style>
        table {
            border-collapse: collapse;
            border: 1px solid black;
            padding: 1px;
            table-layout: fixed;
            width: 655px;
        }

        td {
            padding-top: 3px;
            border-bottom: 1px solid black;
            text-align: center;
            white-space: nowrap;
            border: 1px solid black;
            overflow: hidden;

        }

        th.unit {
            padding: 3px;
            border: 1px solid black;
            background-color: #87CEFA;
            width: 1px;
            font-size: 14px;
            white-space: nowrap;
        }

        td.idcol {
            text-align: center;
            border: 1px solid black;
            width: 1px;
            white-space: nowrap;
            overflow: hidden;
        }
        </style>
    </head>

    <br>
    
        <img src="cid:banner" width="900"> 
        <img src="cid:total_kpi" width="900"> 
        <img src = "cid:img1" width="900"> 

        <table style="width:900px">
          <tr><th colspan='8' style=" background-color: #b2ff66; font-size: 20px; ">All: Sales Achievement % </th></tr>
        
                """ + st.all_Sales_achiv_data() + """</table>
        <br>
        
        <table style="width:900px">
        <tr><th colspan='8' style=" background-color: #b2ff66; font-size: 20px; ">All: Sales Trend % </th></tr>
        
                """ + st.all_Sales_table_data() + """</table>
        <br>
        # <table style="width:900px">
        # <tr><th colspan='8' style=" background-color: #b2ff66; font-size: 20px; ">All: Yesterday Seen Rx </th></tr>
        # 
        
        # <br>
        # 
        # <table style="width:900px">
        # <tr><th colspan='8' style=" background-color: #b2ff66; font-size: 20px; ">All: Yesterday Doctor Call 
        </th></tr>
        # 

        # 
        # </br> </br>
    
        </body>
    
        </html>
        <br><p style="text-align:left;">If there is any inconvenience, you are requested to communicate with the ERP BI Service:<br>
                             <b>(Mobile: 01713-389972, 01713-380502)</b><br><br>
                             <b>Regards</b><br>
                             <b>ERP BI Service</b><br>
                             <b>Information System Automation (ISA)</b><br><br>
                             <p style="color:blue">****This is a system generated report ****</p>
        """, 'html')

    msg.attach(msgText)

    # # Attached Seen Rx File
    # file_location = r"./Data/SeenRx/Seen_Rx_Data.xlsx"
    # filename = os.path.basename(file_location)
    # attachment = open(file_location, "rb")
    # part = MIMEBase('application', 'octet-stream')
    # part.set_payload(attachment.read())
    # encoders.encode_base64(part)
    # part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # msgRoot.attach(part)

    # # Attached Sales Trend File
    # file_location = r"./Data/SalesTrend/sales_trend_data.xlsx"
    # filename = os.path.basename(file_location)
    # attachment = open(file_location, "rb")
    # part = MIMEBase('application', 'octet-stream')
    # part.set_payload(attachment.read())
    # encoders.encode_base64(part)
    # part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # msgRoot.attach(part)

    # # # Attached Doctor Call File
    # file_location = r"./Data/Call/doctor_call_data.xlsx"
    # filename = os.path.basename(file_location)
    # attachment = open(file_location, "rb")
    # part = MIMEBase('application', 'octet-stream')
    # part.set_payload(attachment.read())
    # encoders.encode_base64(part)
    # part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    # msgRoot.attach(part)

    # # ----------- Finally send mail and close server connection ---
    server = smtplib.SMTP(email_server_host, port)
    server.ehlo()
    server.sendmail(me, recipient, msgRoot.as_string())
    server.close()
    print('Mail Send -------- \n')
