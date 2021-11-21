from PIL import Image, ImageDraw, ImageFont
import pytz
import datetime as dd
from PIL import Image
from datetime import datetime


def banner(rsm):
    date = datetime.today()
    day = "Upto" + str(date.day) + '/' + str(date.month) + '/' + str(date.year)
    tz_NY = pytz.timezone('Asia/Dhaka')
    datetime_BD = datetime.now(tz_NY)
    time = datetime_BD.strftime("%I:%M %p")
    date = datetime.today()
    x = dd.datetime.now()
    day = str(date.day) + '-' + str(x.strftime("%b")) + '-' + str(date.year)
    tz_NY = pytz.timezone('Asia/Dhaka')
    datetime_BD = datetime.now(tz_NY)
    time = datetime_BD.strftime("%I:%M %p")

    img = Image.open("./Images/Banner/new_ai.png")
    timestore = ImageDraw.Draw(img)
    branch = ImageDraw.Draw(img)
    font = ImageFont.truetype("./Font/Stencil_Regular.ttf", 35, encoding="unic")
    font1 = ImageFont.truetype("./Font/ROCK.ttf", 35, encoding="unic")
    font2 = ImageFont.truetype("./Font/ROCK.ttf", 18, encoding="unic")
    report_name = ''
    branch.text((25, 130), report_name + "Brand Performance Report - "+str(rsm), (255, 209, 0), font=font1)
    timestore.text((25, 175), "Upto " + day + "," + time, (255, 255, 255), font=font2)
    img.save("./Images/Banner/banner_ai_"+str(rsm)+".png")

    print('1. Banner created')

def all_banner():
    date = datetime.today()
    day = "Upto" + str(date.day) + '/' + str(date.month) + '/' + str(date.year)
    tz_NY = pytz.timezone('Asia/Dhaka')
    datetime_BD = datetime.now(tz_NY)
    time = datetime_BD.strftime("%I:%M %p")
    date = datetime.today()
    x = dd.datetime.now()
    day = str(date.day) + '-' + str(x.strftime("%b")) + '-' + str(date.year)
    tz_NY = pytz.timezone('Asia/Dhaka')
    datetime_BD = datetime.now(tz_NY)
    time = datetime_BD.strftime("%I:%M %p")

    img = Image.open("./Images/Banner/new_ai.png")
    title = ImageDraw.Draw(img)
    timestore = ImageDraw.Draw(img)
    tag = ImageDraw.Draw(img)
    branch = ImageDraw.Draw(img)
    font = ImageFont.truetype("./Font/Stencil_Regular.ttf", 35, encoding="unic")
    font1 = ImageFont.truetype("./Font/ROCK.ttf", 35, encoding="unic")
    font2 = ImageFont.truetype("./Font/ROCK.ttf", 18, encoding="unic")
    report_name = ''
    #
    # tag.text((18, 8), 'SK+F', (255, 255, 255), font=font)
    branch.text((25, 130), report_name + "Brand Performance Report", (255, 209, 0), font=font1)
    timestore.text((25, 175), "Upto " + day + "," + time, (255, 255, 255), font=font2)
    img.save("./Images/Banner/all_banner_ai.png")

    print('Banner created for :All ')
