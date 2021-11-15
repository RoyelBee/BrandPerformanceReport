import xlrd
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from email import encoders
from email.mime.base import MIMEBase
import pandas as pd

import AdditonalFiles.banner as b
import AdditonalFiles.sales_trend_table as st
import AdditonalFiles.seen_rx_table as sr
import AdditonalFiles.doctor_call_table as dc
import AdditonalFiles.all_kpi_image as kpi


def send_report(rsm, email):
    print('---------------------------------\nReport Generating for =', rsm)
    b.banner(rsm)  # # 1

    import AdditonalFiles.all_kpi_image as all_kpi
    all_kpi.create_rsm_total_kpi(rsm)  # # 2
    all_kpi.create_total_kpi()
    kpi.rsm_kpi_image(rsm)
    # ------------ Group email ----------------------------------------
    msgRoot = MIMEMultipart('related')
    me = 'erp-bi.service@transcombd.com'
    # #  --------  For testing purpose mail --------------
    email = 'rejaul.islam@transcombd.com'
    to = [email, '']
    cc = ['', '']
    bcc = ['', '']

    # to = ['', 'tafsir.bashar@skf.transcombd.com']
    # cc = ['', '']
    # bcc = ['yakub@transcombd.com', 'rejaul.islam@transcombd.com']

    # # ------------- Final Mail -----------------
    # to = [email, '']
    # cc = ['khalid@skf.transcombd.com', 'tafsir.bashar@skf.transcombd.com', 'mohammad.nasir@skf.transcombd.com']
    # bcc = ['yakub@transcombd.com', 'tawhid@transcombd.com', 'zubair@transcombd.com', 'rejaul.islam@transcombd.com']

    recipient = to + cc + bcc
    # print('mail Sending to = ', email)

    subject = "United Brand Performance"

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
    img = open("./Images/Banner/banner_ai_" + str(rsm) + ".png", 'rb')
    img1 = MIMEImage(img.read())
    img.close()
    img1.add_header('Content-ID', '<img1>')
    msgRoot.attach(img1)

    # # Total KPI
    total_kpi_img = open("./Images/" + str(rsm) + "_total_kpi_images.png", 'rb')
    total_kpi_img = MIMEImage(total_kpi_img.read())
    total_kpi_img.add_header('Content-ID', '<total_kpi_img>')
    msgRoot.attach(total_kpi_img)

    # # KPI Image
    kpi_img = open("./Images/kpi_image_" + str(rsm) + ".png", 'rb')
    kpi_img = MIMEImage(kpi_img.read())
    kpi_img.add_header('Content-ID', '<kpi_img>')
    msgRoot.attach(kpi_img)

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
        <img src = "cid:img1" width="900"> 
        <img src = "cid:total_kpi_img" width="900"> 
        <img src = "cid:kpi_img" width="900"> 
        
        <table style="width:900px">
          <tr><th colspan='13' style=" background-color: #b2ff66; font-size: 20px; ">All: Sales Achievement % </th></tr>
        
                """ + st.rsm_Sales_achiv_data(rsm) + """</table>
        <br>

        <table style="width:900px">
        <tr><th colspan='13' style=" background-color: #b2ff66; font-size: 20px; ">""" + str(rsm) + """: Sales Trend % 
        </th></tr>
        
               """ + st.rsm_Sales_table_data(rsm) + """</table>
    <br>

        <table style="width:900px">
        <tr><th colspan='13' style=" background-color: #b2ff66; font-size: 20px; ">""" + str(rsm) + """: Yesterday 
        Seen Rx </th></tr>
        
            """ + sr.seen_rx_table_data(rsm) + """</table>
    <br>

        <table style="width:900px">
        <tr><th colspan='13' style=" background-color: #b2ff66; font-size: 20px; ">""" + str(rsm) + """: Yesterday 
        Doctor Call </th></tr>
        """ + dc.doctor_call_table_data(rsm) + """</table>
    
    </br> </br>
    
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
    file_location = r"./Data/SeenRx/SeenRx_" + str(rsm) + ".xlsx"
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msgRoot.attach(part)

    # # Attached Sales Trend File
    file_location = r"./Data/SalesTrend/SalesTrend_" + str(rsm) + ".xlsx"
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msgRoot.attach(part)

    # # Attached Doctor Call File
    file_location = r"./Data/Call/Call_" + str(rsm) + ".xlsx"
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msgRoot.attach(part)

    # # ----------- Finally send mail and close server connection ---
    server = smtplib.SMTP(email_server_host, port)
    server.ehlo()
    server.sendmail(me, recipient, msgRoot.as_string())
    server.close()
    print('Mail Send -------- \n')
