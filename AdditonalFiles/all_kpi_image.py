import pandas as pd
from PIL import ImageDraw, ImageFont
from PIL import Image
import AdditonalFiles.db_connection as conn
from datetime import datetime
from datetime import date, timedelta
import warnings
from statistics import mean
from datetime import datetime
from datetime import date
from calendar import monthrange

warnings.filterwarnings('ignore')

today = datetime.today()
yesterday = date.today() - timedelta(days=1)

day = str(yesterday.day) + '-' + str(yesterday.strftime("%b")) + '-' + str(yesterday.year)

curr_month = str(today.strftime("%b")) + '-' + str(today.year)

font = ImageFont.truetype("./font/ROCK.ttf", 22, encoding="unic")
font1 = ImageFont.truetype("./font/ROCK.ttf", 20, encoding="unic")


def number_decorator(number):
    number = round(number, 2)
    number = format(number, ',')
    return number


def create_total_kpi():
    font = ImageFont.truetype("./font/ROCK.ttf", 40, encoding="unic")
    f = ImageFont.truetype("./font/ROCK.ttf", 25, encoding="unic")

    img = Image.open("./Images/total_kpi.png")
    img_draw = ImageDraw.Draw(img)

    trgdf = pd.read_sql_query("""
             select isnull( sum(amount), 0) as brand_target from V_FF_Brand_TargetAB
            where brand in ( 'Ostocal','Solbion',',Xinc','Maxfer','Zeefol','Danamet','Ethinor','Lenor','Zatral','Urokit','Valenty')  """,
                              conn.azure)

    target = int(trgdf.brand_target)

    salesdf = pd.read_sql_query(""" declare @Rptdate DATE=getdate()-1
                                    declare @DaysInMonth  nvarchar(MAX)
                                    set @DaysInMonth = DAY(EOMONTH(@Rptdate))
                                    
                                    DECLARE @TotalDaysInMonth as Integer=(SELECT DATEDIFF(DAY, @Rptdate, DATEADD(MONTH, 1, @Rptdate)))
                                    DECLARE @totalDaysGone as integer =(SELECT DATEPART(DD, @Rptdate))
                                    
                                    DECLARE @DATE AS SMALLDATETIME = @Rptdate
                                    DECLARE @FIRSTDATEOFMONTH AS SMALLDATETIME = CONVERT(SMALLDATETIME, CONVERT(CHAR(4),YEAR(@DATE)) + '-' + CONVERT(CHAR(2),MONTH(@DATE)) + '-01')
                                    DECLARE @YESTERDAY AS SMALLDATETIME = DATEADD(d,-1,@DATE)
                                    DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), @Rptdate, 112)
                                    DECLARE @FIRSTDATEOFMONTH_STR AS CHAR(8)=CONVERT(VARCHAR(10), @FIRSTDATEOFMONTH , 112)
                                    DECLARE @YESTERDAY_STR AS CHAR(8)=CONVERT(VARCHAR(10), @Rptdate , 112)
                                    DECLARE @BEFOREYESTERDAY_STR AS CHAR(8)=CONVERT(VARCHAR(10), @YESTERDAY , 112)    
                                    select  SUM(dbo.OESalesDetails.EXTINVMISC) AS brandsales 
                                                   
                                    FROM dbo.OESalesDetails RIGHT OUTER JOIN
                                        dbo.PRINFOSKF ON dbo.OESalesDetails.ITEM = dbo.PRINFOSKF.ITEMNO
                                    WHERE dbo.PRINFOSKF.BRAND IN ('Ostocal', 'Solbion',',Xinc','Maxfer','Zeefol','Danamet','Ethinor','Lenor','Zatral','Urokit','Valenty')
                                    and (dbo.OESalesDetails.TRANSDATE BETWEEN @FIRSTDATEOFMONTH_STR AND @YESTERDAY_STR) 
                  """, conn.azure)
    sales = int(salesdf.brandsales)

    if target > 1:
        sales_achivPercent = round((sales / target) * 100, 2)
    else:
        sales_achivPercent = 0
    #
    today = str(date.today())
    datee = datetime.strptime(today, "%Y-%m-%d")
    date_val = int(datee.day)
    num_days = monthrange(datee.year, datee.month)[1]

    if target > 1:
        trend_val = (sales / date_val) * num_days
        trend_percent = round((trend_val / target) * 100, 2)
    else:
        trend_percent = 0

    # # # ------------------------- Seen Rx -----------------------------------
    seen_df = pd.read_csv('./Data/SeenRx/rsm_seen_total.csv')
    OSTOCAL_seen = sum(seen_df.OSTOCAL.tolist())
    SOLBION_seen = sum(seen_df.SOLBION.tolist())
    XINC_seen = sum(seen_df.XINC.tolist())
    MAXFER_seen = sum(seen_df.MAXFER.tolist())
    ZEEFOL_seen = sum(seen_df.ZEEFOL.tolist())
    ETHINOR_seen = sum(seen_df.ETHINOR.tolist())
    LENOR_seen = sum(seen_df.LENOR.tolist())
    ZATRAL_seen = sum(seen_df.ZATRAL.tolist())
    UROKIT_seen = sum(seen_df.UROKIT.tolist())
    VALENTY_seen = sum(seen_df.VALENTY.tolist())
    all_seenrx = int(
        sum([OSTOCAL_seen, SOLBION_seen, XINC_seen, MAXFER_seen, ZEEFOL_seen, ETHINOR_seen, LENOR_seen, ZATRAL_seen,
             UROKIT_seen, VALENTY_seen]))

    # # ---------------------- Doctor Call -----------------------------
    call_df = pd.read_csv('./Data/Call/rsm_call_total.csv')
    OSTOCAL_call = sum(call_df.OSTOCAL.tolist())
    SOLBION_call = sum(call_df.SOLBION.tolist())
    XINC_call = sum(call_df.XINC.tolist())
    MAXFER_call = sum(seen_df.MAXFER.tolist())
    ZEEFOL_call = sum(seen_df.ZEEFOL.tolist())
    ETHINOR_call = sum(seen_df.ETHINOR.tolist())
    LENOR_call = sum(seen_df.LENOR.tolist())
    ZATRAL_call = sum(seen_df.ZATRAL.tolist())
    UROKIT_call = sum(seen_df.UROKIT.tolist())
    VALENTY_call = sum(seen_df.VALENTY.tolist())

    all_call = int(
        sum([OSTOCAL_call, SOLBION_call, XINC_call, MAXFER_call, ZEEFOL_call, ETHINOR_call, ZATRAL_call, VALENTY_call,
             LENOR_call, UROKIT_call]))

    img_draw.text((130, 10), '(' + str(day) + ')', (255, 255, 255), f)
    img_draw.text((100, 110), number_decorator(sales_achivPercent) + '%', (0, 0, 0), font)
    img_draw.text((420, 110), number_decorator(trend_percent) + '%', (0, 0, 0), font)
    img_draw.text((730, 110), number_decorator(all_seenrx), (0, 0, 0), font)
    img_draw.text((1030, 110), number_decorator(all_call), (0, 0, 0), font)

    img.save("./Images/total_kpi_images.png")
    print('3. Specefic KPI Created')


# def create_rsm_total_kpi(rsm):
#     font = ImageFont.truetype("./font/ROCK.ttf", 40, encoding="unic")
#     f = ImageFont.truetype("./font/ROCK.ttf", 25, encoding="unic")
#
#     img = Image.open("./Images/total_kpi.png")
#     img_draw = ImageDraw.Draw(img)
#
#     trgdf = pd.read_sql_query(""" select sum(amount) as brand_target from V_FF_Brand_Target
#             where rsmtr like ? and brand in ( 'LIGAZID', 'EMAZID', 'LIPICON', 'AGLIP', 'CIFIBET',
#                             'AMLEVO', 'CARDOBIS', 'RIVAROX', 'NOCLOG', 'BEMPID',
#                             'AROTIDE', 'FOBUNID') """, conn.azure, params={rsm})
#
#     target = int(trgdf.brand_target)
#
#     salesdf = pd.read_sql_query(""" DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
#             DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE()-1, 112)
#             select sum(extinvmisc) as brandsales from V_FF_Brand_Sales where left(MSOTR, 3) like ? and left(transdate,
#             6)=@This_month and  TRANSDATE<=@LastDay and
#             brand in ( 'LIGAZID', 'EMAZID', 'LIPICON', 'AGLIP', 'CIFIBET',
#             'AMLEVO', 'CARDOBIS', 'RIVAROX', 'NOCLOG', 'BEMPID',
#             'AROTIDE', 'FOBUNID')   """, conn.azure, params={rsm})
#     sales = int(salesdf.brandsales)
#
#     if target > 1:
#         sales_achivPercent = round((sales / target) * 100, 2)
#     else:
#         sales_achivPercent = 0
#     #
#     today = str(date.today())
#     datee = datetime.strptime(today, "%Y-%m-%d")
#     date_val = int(datee.day)
#     num_days = monthrange(datee.year, datee.month)[1]
#
#     if target > 1:
#         trend_val = (sales / date_val) * num_days
#         trend_percent = round((trend_val / target) * 100, 2)
#     else:
#         trend_percent = 0
#
#     # # # ------------------------- Seen Rx -----------------------------------
#     seen_df = pd.read_csv('./Data/SeenRx/rsm_seen_total.csv')
#     seen_df.where(seen_df['FFTR'] == rsm, inplace=True)
#     total_seenrx = int(seen_df[seen_df['FFTR'].notnull()].sum(axis=1))
#
#     # # # ---------------------- Doctor Call -----------------------------
#     call_df = pd.read_csv('./Data/Call/rsm_call_total.csv')
#     call_df.where(call_df['FFTR'] == rsm, inplace=True)
#     total_call = int(call_df[call_df['FFTR'].notnull()].sum(axis=1))
#
#     #
#     img_draw.text((130, 22), '(' + str(day) + ')', (255, 255, 255), f)
#     img_draw.text((90, 120), number_decorator(sales_achivPercent) + '%', (0, 0, 0), font)
#     img_draw.text((410, 120), number_decorator(trend_percent) + '%', (0, 0, 0), font)
#     img_draw.text((750, 90), number_decorator(total_seenrx), (0, 0, 0), font)
#     img_draw.text((1050, 120), number_decorator(total_call), (0, 0, 0), font)
#
#     img.save("./Images/" + str(rsm) + "_total_kpi_images.png")
#     print('2. Total Summary KPI Generated')


def all_kpi_images():
    img1 = Image.open("./Images/kpi_image.png")
    img_draw = ImageDraw.Draw(img1)

    sales_achiv = pd.read_excel('./Data/SalesAchivment/sales_achiv.xlsx')
    if sales_achiv.empty == True:
        OSTOCAL_achiv = SOLBION_achiv = XINC_achiv = MAXFER_achiv = ZEEFOL_achiv = DANAMET_achiv = 0
        ETHINOR_achiv = LENOR_achiv = ZATRAL_achiv = UROKIT_achiv = VALENTY_achiv = 0

    else:
        OSTOCAL_achiv = sales_achiv.OSTOCAL.tolist()[0]
        SOLBION_achiv = sales_achiv.SOLBION.tolist()[0]
        XINC_achiv = sales_achiv.XINC.tolist()[0]
        MAXFER_achiv = sales_achiv.MAXFER.tolist()[0]
        DANAMET_achiv = sales_achiv.DANAMET.tolist()[0]
        ZEEFOL_achiv = sales_achiv.ZEEFOL.tolist()[0]
        ETHINOR_achiv = sales_achiv.ETHINOR.tolist()[0]
        LENOR_achiv = sales_achiv.LENOR.tolist()[0]
        ZATRAL_achiv = sales_achiv.ZATRAL.tolist()[0]
        UROKIT_achiv = sales_achiv.UROKIT.tolist()[0]
        VALENTY_achiv = sales_achiv.VALENTY.tolist()[0]

    img_draw.text((265, 5), 'up to (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((45, 75), number_decorator(OSTOCAL_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((150, 75), number_decorator(SOLBION_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((250, 75), number_decorator(XINC_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((350, 75), number_decorator(MAXFER_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((450, 75), number_decorator(ZEEFOL_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((560, 75), number_decorator(DANAMET_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((660, 75), number_decorator(ETHINOR_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((760, 75), number_decorator(LENOR_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((870, 75), number_decorator(ZATRAL_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((960, 75), number_decorator(UROKIT_achiv) + '%', (0, 0, 0), font1)
    img_draw.text((1070, 75), number_decorator(VALENTY_achiv) + '%', (0, 0, 0), font1)
    # ------------------------------------------------------------------------------
    sales_trend = pd.read_excel('./Data/SalesTrend/sales_trend_data.xlsx')
    if sales_trend.empty == True:
        OSTOCAL_trend = SOLBION_trend = XINC_trend = MAXFER_trend = ZEEFOL_trend = DANAMET_trend = 0
        ETHINOR_trend = LENOR_trend = ZATRAL_trend = UROKIT_trend = VALENTY_trend = 0

    else:
        OSTOCAL_trend = sales_trend.OSTOCAL.tolist()[0]
        SOLBION_trend = sales_trend.SOLBION.tolist()[0]
        XINC_trend = sales_trend.XINC.tolist()[0]
        MAXFER_trend = sales_achiv.MAXFER.tolist()[0]
        ZEEFOL_trend = sales_achiv.ZEEFOL.tolist()[0]
        DANAMET_trend = sales_achiv.DANAMET.tolist()[0]
        ETHINOR_trend = sales_achiv.ETHINOR.tolist()[0]
        LENOR_trend = sales_achiv.LENOR.tolist()[0]
        ZATRAL_trend = sales_achiv.ZATRAL.tolist()[0]
        UROKIT_trend = sales_achiv.UROKIT.tolist()[0]
        VALENTY_trend = sales_achiv.VALENTY.tolist()[0]

    img_draw.text((180, 120), 'up to (' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((45, 195), number_decorator(OSTOCAL_trend) + '%', (0, 0, 0), font1)
    img_draw.text((150, 195), number_decorator(SOLBION_trend) + '%', (0, 0, 0), font1)
    img_draw.text((250, 195), number_decorator(XINC_trend) + '%', (0, 0, 0), font1)
    img_draw.text((350, 195), number_decorator(MAXFER_trend) + '%', (0, 0, 0), font1)
    img_draw.text((450, 195), number_decorator(ZEEFOL_trend) + '%', (0, 0, 0), font1)
    img_draw.text((560, 195), number_decorator(DANAMET_trend) + '%', (0, 0, 0), font1)
    img_draw.text((660, 195), number_decorator(ETHINOR_trend) + '%', (0, 0, 0), font1)
    img_draw.text((760, 195), number_decorator(LENOR_trend) + '%', (0, 0, 0), font1)
    img_draw.text((870, 195), number_decorator(ZATRAL_trend) + '%', (0, 0, 0), font1)
    img_draw.text((960, 195), number_decorator(UROKIT_trend) + '%', (0, 0, 0), font1)
    img_draw.text((1070, 195), number_decorator(VALENTY_trend) + '%', (0, 0, 0), font1)

    # # ------------------------------ Seen Rx ----------------------------------------
    seen_rx = pd.read_csv('./Data/SeenRx/rsm_seen_total.csv')
    OSTOCAL_seen = sum(seen_rx.OSTOCAL.tolist())
    SOLBION_seen = sum(seen_rx.SOLBION.tolist())
    XINC_seen = sum(seen_rx.XINC.tolist())
    MAXFER_seen = sum(seen_rx.MAXFER.tolist())
    DANMET_seen = sum(seen_rx.DANMET.tolist())
    ZEEFOL_seen = sum(seen_rx.ZEEFOL.tolist())
    ETHINOR_seen = sum(seen_rx.ETHINOR.tolist())
    LENOR_seen = sum(seen_rx.LENOR.tolist())
    ZATRAL_seen = sum(seen_rx.ZATRAL.tolist())
    UROKIT_seen = sum(seen_rx.UROKIT.tolist())
    VALENTY_seen = sum(seen_rx.VALENTY.tolist())
    # print(VALENTY_seen)

    img_draw.text((117, 243), '(' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((45, 315), number_decorator(OSTOCAL_seen), (0, 0, 0), font1)
    img_draw.text((150, 315), number_decorator(SOLBION_seen), (0, 0, 0), font1)
    img_draw.text((250, 315), number_decorator(XINC_seen), (0, 0, 0), font1)
    img_draw.text((350, 315), number_decorator(MAXFER_seen), (0, 0, 0), font1)
    img_draw.text((450, 315), number_decorator(ZEEFOL_seen), (0, 0, 0), font1)
    img_draw.text((560, 315), number_decorator(DANMET_seen), (0, 0, 0), font1)
    img_draw.text((660, 315), number_decorator(ETHINOR_seen), (0, 0, 0), font1)
    img_draw.text((760, 315), number_decorator(LENOR_seen), (0, 0, 0), font1)
    img_draw.text((870, 315), number_decorator(ZATRAL_seen), (0, 0, 0), font1)
    img_draw.text((960, 315), number_decorator(UROKIT_seen), (0, 0, 0), font1)
    img_draw.text((1070, 315), number_decorator(VALENTY_seen), (0, 0, 0), font1)

    # # # ------------------------------ Doctor Call -----------------------------------
    doctor_call = pd.read_csv('./Data/Call/rsm_call_total.csv')
    OSTOCAL_call = sum(doctor_call.OSTOCAL.tolist())
    SOLBION_call = sum(doctor_call.SOLBION.tolist())
    XINC_call = sum(doctor_call.XINC.tolist())
    MAXFER_call = sum(doctor_call['MAXFER Call'].tolist())
    DANAMET_call = sum(doctor_call['DANAMET Call'].tolist())
    ZEEFOL_call = sum(doctor_call['ZEEFOL Call'].tolist())
    ETHINOR_call = sum(doctor_call['ETHINOR Call'].tolist())
    LENOR_call = sum(doctor_call['LENOR Call'].tolist())
    ZATRAL_call = sum(doctor_call['ZATRAL Call'].tolist())
    UROKIT_call = sum(doctor_call['UROKIT Call'].tolist())
    VALENTY_call = sum(doctor_call['VALENTY Call'].tolist())

    #
    img_draw.text((140, 368), '(' + str(day) + ')', (255, 255, 255), font)
    img_draw.text((27, 445), number_decorator(OSTOCAL_call), (0, 0, 0), font1)
    img_draw.text((130, 445), number_decorator(SOLBION_call), (0, 0, 0), font1)
    img_draw.text((230, 445), number_decorator(XINC_call), (0, 0, 0), font1)
    img_draw.text((350, 445), number_decorator(MAXFER_call), (0, 0, 0), font1)
    img_draw.text((450, 445), number_decorator(ZEEFOL_call), (0, 0, 0), font1)
    img_draw.text((560, 445), number_decorator(DANAMET_call), (0, 0, 0), font1)
    img_draw.text((660, 445), number_decorator(ETHINOR_call), (0, 0, 0), font1)
    img_draw.text((760, 445), number_decorator(LENOR_call), (0, 0, 0), font1)
    img_draw.text((870, 445), number_decorator(ZATRAL_call), (0, 0, 0), font1)
    img_draw.text((960, 445), number_decorator(UROKIT_call), (0, 0, 0), font1)
    img_draw.text((1070, 445), number_decorator(VALENTY_call), (0, 0, 0), font1)
    img1.save("./Images/all_kpi_image.png")
    print("All KPI Image created")


def rsm_kpi_image(rsm):
    img1 = Image.open("./Images/kpi.png")
    img_draw = ImageDraw.Draw(img1)

    sales_query = (""" DECLARE @YearMonth as VARCHAR(6) = Convert (varchar,Getdate(),112)
        DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE(), 112)
        DECLARE @FIRSTDATEOFMONTH AS CHAR(8) = CONVERT(VARCHAR(6), GETDATE(), 112)+'01'
        DECLARE @Today as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)
        DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
        DECLARE @TotalDaysInMonth as Integer=(SELECT DATEDIFF(DAY, getdate(), DATEADD(MONTH, 1, getdate())))
        DECLARE @TotalDaysGone as integer =(select  count(distinct transdate) from OESalesSummery where left(transdate,6)=@This_month and  TRANSDATE<=@LastDay)

        Select
        isnull(Sum(Case when FFTarget.Brand = 'LIGAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS LIGAZID,
        isnull(Sum(Case when FFTarget.Brand = 'EMAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS EMAZID,
        isnull(Sum(Case when FFTarget.Brand = 'LIPICON' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS LIPICON,
        isnull(Sum(Case when FFTarget.Brand = 'AGLIP' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS AGLIP,
        isnull(Sum(Case when FFTarget.Brand = 'CIFIBET' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS CIFIBET,
        isnull(Sum(Case when FFTarget.Brand = 'AMLEVO' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS AMLEVO,
        isnull(Sum(Case when FFTarget.Brand = 'CARDOBIS' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS CARDOBIS,
        isnull(Sum(Case when FFTarget.Brand = 'RIVAROX' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS RIVAROX,
        isnull(Sum(Case when FFTarget.Brand = 'NOCLOG' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0) AS NOCLOG

        from
        (Select BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target group by BRAND) as FFTarget
         left join
        (Select BRAND, SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_Sales where TRANSDATE <= @LastDay  group by BRAND) as FFSales
        ON  (FFTarget.BRAND=FFSales.BRAND)
    """)
    sales_data = pd.read_sql_query(sales_query, conn.azure)
    sales_data = pd.read_excel("./Data/SalesTrend/SalesTrend_" + str(rsm) + ".xlsx")
    if sales_data.empty == True:
        LIGZID_data = EMAZID_data = LIPICON_data = AGLIP_data = CIFIBET_data = AMLEVO_data = CARDOBIS_data = \
            RIVAROX_data = NOCLOG_data = 0
    else:
        LIGZID_data = sales_data.LIGAZID.tolist()[0]
        EMAZID_data = sales_data.EMAZID.tolist()[0]
        LIPICON_data = sales_data.LIPICON.tolist()[0]
        AGLIP_data = sales_data.AGLIP.tolist()[0]
        CIFIBET_data = sales_data.CIFIBET.tolist()[0]
        AMLEVO_data = sales_data.AMLEVO.tolist()[0]
        CARDOBIS_data = sales_data.CARDOBIS.tolist()[0]
        RIVAROX_data = sales_data.RIVAROX.tolist()[0]
        NOCLOG_data = sales_data.NOCLOG.tolist()[0]

    img_draw.text((180, 5), '(' + str(day) + ') ' + str(rsm), (255, 255, 255), font)
    img_draw.text((25, 95), number_decorator(LIGZID_data) + '%', (0, 0, 0), font1)
    img_draw.text((160, 95), number_decorator(EMAZID_data) + '%', (0, 0, 0), font1)
    img_draw.text((290, 95), number_decorator(LIPICON_data) + '%', (0, 0, 0), font1)
    img_draw.text((425, 95), number_decorator(AGLIP_data) + '%', (0, 0, 0), font1)
    img_draw.text((555, 95), number_decorator(CIFIBET_data) + '%', (0, 0, 0), font1)
    img_draw.text((690, 95), number_decorator(AMLEVO_data) + '%', (0, 0, 0), font1)
    img_draw.text((820, 95), number_decorator(CARDOBIS_data) + '%', (0, 0, 0), font1)
    img_draw.text((950, 95), number_decorator(RIVAROX_data) + '%', (0, 0, 0), font1)
    img_draw.text((1080, 95), number_decorator(NOCLOG_data) + '%', (0, 0, 0), font1)

    seenrx_query = ("""
                select * from
                (Select
                left([FF ID],3) as [FF ID],
                isnull(sum([LIGAZID Seen Rx]),0)  as [LIGAZID],
                isnull(sum([EMAZID Seen Rx]),0) as [EMAZID],
                isnull(sum([LIPICON Seen Rx]),0) as [LIPICON],
                isnull(sum([AGLIP Seen Rx]),0) as [AGLIP],
                isnull(sum([CIFIBET Seen Rx]),0) as [CIFIBET],
                isnull(sum([AMLEVO Seen Rx]),0) as [AMLEVO],
                isnull(sum([CARDOBIS Seen Rx]),0) as [CARDOBIS],
                isnull(sum([RIVAROX Seen Rx]),0) as [RIVAROX],
                isnull(sum([NOCLOG Seen Rx]),0) as [NOCLOG]
                from v_LastDay_SeenRx
                where [FF ID]  = left([FF ID],4)+'0' and left([FF ID],3) like ?
                group by left([FF ID],3)

                union all

                Select [FF ID],
                isnull([LIGAZID Seen Rx],0) as [LIGAZID Seen Rx],
                isnull([EMAZID Seen Rx],0) as [EMAZID Seen Rx],
                ISNULL([LIPICON Seen Rx], 0) as [LIPICON Seen Rx],
                ISNULL([AGLIP Seen Rx], 0) as [AGLIP Seen Rx],
                ISNULL([CIFIBET Seen Rx], 0) as [CIFIBET Seen Rx],
                ISNULL([AMLEVO Seen Rx], 0) as [AMLEVO Seen Rx],
                ISNULL([CARDOBIS Seen Rx], 0) as [CARDOBIS Seen Rx],
                ISNULL([RIVAROX Seen Rx], 0) as [RIVAROX Seen Rx],
                ISNULL([NOCLOG Seen Rx], 0) as [NOCLOG Seen Rx]
                from v_LastDay_SeenRx where  left([FF ID],3) like ?

                ) as T1
                order by [FF ID] asc
            """)

    seenrx_seen_data = pd.read_sql_query(seenrx_query, conn.m_reporting, params=(rsm, rsm))

    LIGZID_seen_seen_data = seenrx_seen_data.LIGAZID.tolist()[0]
    EMAZID_seen_data = seenrx_seen_data.EMAZID.tolist()[0]
    LIPICON_seen_data = seenrx_seen_data.LIPICON.tolist()[0]
    AGLIP_seen_data = seenrx_seen_data.AGLIP.tolist()[0]
    CIFIBET_seen_data = seenrx_seen_data.CIFIBET.tolist()[0]
    AMLEVO_seen_data = seenrx_seen_data.AMLEVO.tolist()[0]
    CARDOBIS_seen_data = seenrx_seen_data.CARDOBIS.tolist()[0]
    RIVAROX_seen_data = seenrx_seen_data.RIVAROX.tolist()[0]
    NOCLOG_seen_data = seenrx_seen_data.NOCLOG.tolist()[0]

    # # Seen Rx
    img_draw.text((120, 152), '(' + str(day) + ') ' + str(rsm), (255, 255, 255), font)
    img_draw.text((55, 235), number_decorator(LIGZID_seen_seen_data), (0, 0, 0), font1)
    img_draw.text((190, 235), number_decorator(EMAZID_seen_data), (0, 0, 0), font1)
    img_draw.text((310, 235), number_decorator(LIPICON_seen_data), (0, 0, 0), font1)
    img_draw.text((455, 235), number_decorator(AGLIP_seen_data), (0, 0, 0), font1)
    img_draw.text((590, 235), number_decorator(CIFIBET_seen_data), (0, 0, 0), font1)
    img_draw.text((730, 235), number_decorator(AMLEVO_seen_data), (0, 0, 0), font1)
    img_draw.text((860, 235), number_decorator(CARDOBIS_seen_data), (0, 0, 0), font1)
    img_draw.text((990, 235), number_decorator(RIVAROX_seen_data), (0, 0, 0), font1)
    img_draw.text((1110, 235), number_decorator(NOCLOG_seen_data), (0, 0, 0), font1)

    call_data = pd.read_excel("./Data/Call/call_" + str(rsm) + ".xlsx")
    LIGAZID_Call = sum(call_data['LIGAZID'])
    EMAZID_Call = sum(call_data['EMAZID'])
    LIPICON_Call = sum(call_data['LIPICON'])
    AGLIP_Call = sum(call_data['AGLIP'])
    CIFIBET_Call = sum(call_data['CIFIBET'])
    AMLEVO_Call = sum(call_data['AMLEVO'])
    CARDOBIS_Call = sum(call_data['CARDOBIS'])
    RIVAROX_Call = sum(call_data['RIVAROX'])
    Noclog_Call = sum(call_data['NOCLOG'])

    # # Doctor Call
    img_draw.text((155, 300), '(' + str(day) + ') ' + str(rsm), (255, 255, 255), font)
    img_draw.text((50, 390), number_decorator(LIGAZID_Call), (0, 0, 0), font1)
    img_draw.text((180, 390), number_decorator(EMAZID_Call), (0, 0, 0), font1)
    img_draw.text((310, 390), number_decorator(LIPICON_Call), (0, 0, 0), font1)
    img_draw.text((455, 390), number_decorator(AGLIP_Call), (0, 0, 0), font1)
    img_draw.text((580, 390), number_decorator(CIFIBET_Call), (0, 0, 0), font1)
    img_draw.text((720, 390), number_decorator(AMLEVO_Call), (0, 0, 0), font1)
    img_draw.text((845, 390), number_decorator(CARDOBIS_Call), (0, 0, 0), font1)
    img_draw.text((990, 390), number_decorator(RIVAROX_Call), (0, 0, 0), font1)
    img_draw.text((1110, 390), number_decorator(Noclog_Call), (0, 0, 0), font1)
    img1.save("./Images/kpi_image_" + str(rsm) + ".png")
    print('3. Speceric KPI Image Generated')
