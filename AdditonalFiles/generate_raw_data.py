import pandas as pd
import AdditonalFiles.db_connection as conn


def sales_achiv_trend_data():
    sales_achiv_trend_df = pd.read_sql_query("""
        DECLARE @YearMonth as VARCHAR(6) = Convert (varchar,Getdate()-1,112)
        DECLARE @This_month as CHAR(6)= CONVERT(VARCHAR(6), GETDATE()-1, 112)
        DECLARE @FIRSTDATEOFMONTH AS CHAR(8) = CONVERT(VARCHAR(6), GETDATE()-1, 112)+'01'
        DECLARE @Today as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)
        DECLARE @LastDay as CHAR(8) = CONVERT(VARCHAR(8), GETDATE(), 112)-1
        DECLARE @TotalDaysInMonth as Integer=(SELECT DATEDIFF(DAY, getdate()-1, DATEADD(MONTH, 1, getdate())))
        DECLARE @TotalDaysGone as integer =(select  count(distinct transdate) from OESalesSummery where left(transdate,6)=@This_month and  TRANSDATE<=@LastDay)
        
        Select * from
        (Select
        FFTarget.FFTR,
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.BRAND = 'OSTOCAL' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [OSTOCAL SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'OSTOCAL' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [OSTOCAL TREND%],
                
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'SOLBION' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [SOLBION SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'SOLBION' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [SOLBION TREND%],
                
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'XINC' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [XINC SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'XINC' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [XINC TREND%],
                        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Maxfer' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Maxfer SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Maxfer' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Maxfer TREND%], 
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Zeefol' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Zeefol SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Zeefol' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Zeefol TREND%], 
                        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Danamet' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Danamet SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Danamet' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Danamet TREND%], 
                        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Ethinor' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Ethinor SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Ethinor' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Ethinor TREND%], 
                        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Lenor' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Lenor SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Lenor' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Lenor TREND%], 
        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Zatral' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Zatral SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Zatral' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Zatral TREND%], 
        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Urokit' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Urokit SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Urokit' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Urokit TREND%], 
        
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Valenty' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [Valenty SALES ACHIV%],
        Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Valenty' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [Valenty TREND%]
                        
        from(
        Select 'RSM' AS FF, RSMTR as FFTR,BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_TargetAB
        group by RSMTR,BRAND
        union all
        Select 'FM' AS FF, FMTR as FFTR,BRAND,SUM(Amount) AS TargetAmount  from V_FF_Brand_TargetAB
        group by FMTR,BRAND
        union all
        Select 'MSO' AS FF, MSOTR as FFTR,BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_TargetAB
        group by MSOTR,BRAND
        ) as FFTarget
        left join
        (
        Select 'RSM' AS FF, left(MSOTR,2)  as FFTR,BRAND, SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_SalesAB
        where TRANSDATE <= @LastDay
        group by left(MSOTR,2) ,BRAND
        union all
        Select 'FM' AS FF, FMTR as FFTR,BRAND,SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_SalesAB
        where TRANSDATE <= @LastDay
        group by FMTR ,BRAND
        union all
        Select 'MSO' AS FF, MSOTR as FFTR,BRAND, SUM(extinvmisc) AS SalesAmount   from V_FF_Brand_SalesAB
        where TRANSDATE <= @LastDay
        group by MSOTR,BRAND
        ) as FFSales
        ON (FFTarget.FFTR=FFSales.FFTR) AND (FFTarget.BRAND=FFSales.BRAND)
        WHERE FFTarget.FFTR IS NOT NULL
        group by FFTarget.FFTR
        ) as T1
        order by FFTR asc    
            """, conn.azure)

    sales_achiv_trend_df.to_excel('./Data/SalesTrend/sales_achiv_trend_data.xlsx', index=False)

    achivment_col = ['FFTR', 'OSTOCAL SALES ACHIV%', 'SOLBION SALES ACHIV%',
                     'XINC SALES ACHIV%', 'Maxfer SALES ACHIV%', 'Zeefol SALES ACHIV%',
                     'Danamet SALES ACHIV%', 'Ethinor SALES ACHIV%', 'Lenor SALES ACHIV%',
                     'Zatral SALES ACHIV%', 'Urokit SALES ACHIV%', 'Valenty SALES ACHIV%',
                     ]

    trend_col = ['FFTR', 'OSTOCAL TREND%', 'SOLBION TREND%', 'XINC TREND%',
                 'Maxfer TREND%', 'Zeefol TREND%', 'Danamet TREND%', 'Ethinor TREND%',
                 'Lenor TREND%', 'Zatral TREND%', 'Urokit TREND%',
                 'Valenty TREND%'
                 ]

    sales_achiv = pd.read_excel('./Data/SalesTrend/sales_achiv_trend_data.xlsx', usecols=achivment_col)
    sales_achiv.rename(columns={'FFTR': 'FFTR',
                                'OSTOCAL SALES ACHIV%': 'OSTOCAL',
                                'SOLBION SALES ACHIV%': 'SOLBION',
                                'XINC SALES ACHIV%': 'XINC',
                                'Maxfer SALES ACHIV%': 'MAXFER',
                                'Zeefol SALES ACHIV%': 'ZEEFOL',
                                'Danamet SALES ACHIV%': 'DANAMET',
                                'Ethinor SALES ACHIV%': 'ETHINOR',
                                'Lenor SALES ACHIV%': 'LENOR',
                                'Zatral SALES ACHIV%': 'ZATRAL',
                                'Urokit SALES ACHIV%': 'UROKIT',
                                'Valenty SALES ACHIV%': 'VALENTY'

                                }, inplace=True)

    trend_achiv = pd.read_excel('./Data/SalesTrend/sales_achiv_trend_data.xlsx', usecols=trend_col)
    trend_achiv.rename(columns={'FFTR': 'FFTR',
                                'OSTOCAL TREND%': 'OSTOCAL',
                                'SOLBION TREND%': 'SOLBION',
                                'XINC TREND%': 'XINC',
                                'Maxfer TREND%': 'MAXFER',
                                'Zeefol TREND%': 'ZEEFOL',
                                'Danamet TREND%': 'DANAMET',
                                'Ethinor TREND%': 'ETHINOR',
                                'Lenor TREND%': 'LENOR',
                                'Zatral TREND%': 'ZATRAL',
                                'Urokit TREND%': 'UROKIT',
                                'Valenty TREND%': 'VALENTY'
                                }, inplace=True)

    sales_achiv.to_excel("./Data/SalesAchivment/sales_achiv.xlsx", index=False)
    trend_achiv.to_excel("./Data/SalesTrend/sales_trend_data.xlsx", index=False)

    # ## Seperate all achivement data by rsm
    # salesachivper = pd.read_excel('./Data/SalesAchivment/sales_achiv.xlsx')
    # CBUs = salesachivper[salesachivper['FFTR'].str.contains('CBU')]
    # CBUs.to_excel('./Data/SalesAchivment/SalesAchivment_CBU.xlsx', index=False)
    #
    # CCFs = salesachivper[salesachivper['FFTR'].str.contains('CCF')]
    # CCFs.to_excel('./Data/SalesAchivment/SalesAchivment_CCF.xlsx', index=False)
    #
    # CCXs = salesachivper[salesachivper['FFTR'].str.contains('CCX')]
    # CCXs.to_excel('./Data/SalesAchivment/SalesAchivment_CCX.xlsx', index=False)
    #
    # CNHs = salesachivper[salesachivper['FFTR'].str.contains('CNH')]
    # CNHs.to_excel('./Data/SalesAchivment/SalesAchivment_CNH.xlsx', index=False)
    #
    # CKJs = salesachivper[salesachivper['FFTR'].str.contains('CKJ')]
    # CKJs.to_excel('./Data/SalesAchivment/SalesAchivment_CKJ.xlsx', index=False)
    #
    # CMTs = salesachivper[salesachivper['FFTR'].str.contains('CMT')]
    # CMTs.to_excel('./Data/SalesAchivment/SalesAchivment_CMT.xlsx', index=False)
    #
    # CRBs = salesachivper[salesachivper['FFTR'].str.contains('CRB')]
    # CRBs.to_excel('./Data/SalesAchivment/SalesAchivment_CRB.xlsx', index=False)
    #
    # CRPs = salesachivper[salesachivper['FFTR'].str.contains('CRP')]
    # CRPs.to_excel('./Data/SalesAchivment/SalesAchivment_CRP.xlsx', index=False)
    #
    # CSBs = salesachivper[salesachivper['FFTR'].str.contains('CSB')]
    # CSBs.to_excel('./Data/SalesAchivment/SalesAchivment_CSB.xlsx', index=False)
    #
    # ## Seperate all trend data by rsm
    # data = pd.read_excel('./Data/SalesTrend/sales_trend_data.xlsx')
    # CBU = data[data['FFTR'].str.contains('CBU')]
    # CBU.to_excel('./Data/SalesTrend/SalesTrend_CBU.xlsx', index=False)
    #
    # CCF = data[data['FFTR'].str.contains('CCF')]
    # CCF.to_excel('./Data/SalesTrend/SalesTrend_CCF.xlsx', index=False)
    #
    # CCX = data[data['FFTR'].str.contains('CCX')]
    # CCX.to_excel('./Data/SalesTrend/SalesTrend_CCX.xlsx', index=False)
    #
    # CNH = data[data['FFTR'].str.contains('CNH')]
    # CNH.to_excel('./Data/SalesTrend/SalesTrend_CNH.xlsx', index=False)
    #
    # CKJ = data[data['FFTR'].str.contains('CKJ')]
    # CKJ.to_excel('./Data/SalesTrend/SalesTrend_CKJ.xlsx', index=False)
    #
    # CMT = data[data['FFTR'].str.contains('CMT')]
    # CMT.to_excel('./Data/SalesTrend/SalesTrend_CMT.xlsx', index=False)
    #
    # CRB = data[data['FFTR'].str.contains('CRB')]
    # CRB.to_excel('./Data/SalesTrend/SalesTrend_CRB.xlsx', index=False)
    #
    # CRP = data[data['FFTR'].str.contains('CRP')]
    # CRP.to_excel('./Data/SalesTrend/SalesTrend_CRP.xlsx', index=False)
    #
    # CSB = data[data['FFTR'].str.contains('CSB')]
    # CSB.to_excel('./Data/SalesTrend/SalesTrend_CSB.xlsx', index=False)
    # #
    print('1. Sales Achiv and Trend Data Saved')


def seen_rx_data():
    new_seen_rx_data = pd.read_sql_query("""     
        select * from
        (Select
        left([FF ID],2) as [FFTR],
        isnull(sum([OSTOCAL Seen Rx]),0)  as [OSTOCAL],
        isnull(sum([SOLBION Seen Rx]),0)  as [SOLBION],
        isnull(sum([XINC Seen Rx]),0)  as [XINC],
        isnull(SUM([Maxfer Seen Rx]),0) as MAXFER,
		isnull(SUM([Zeefol Seen Rx]),0) as ZEEFOL,
		isnull(SUM([Danamet Seen Rx]),0) as DANMET,
		isnull(SUM([Ethinor Seen Rx]),0) as ETHINOR,
		isnull(SUM([Lenor Seen Rx]),0) as LENOR,
		isnull(SUM([Zatral Seen Rx]),0) as ZATRAL,
		isnull(SUM([Urokit Seen Rx]),0) as UROKIT,
		isnull(SUM([Valenty Seen Rx]),0) as VALENTY
        from v_LastDay_SeenRxAB
        group by left([FF ID],2)
        union all
        Select [FF ID],
        isnull([OSTOCAL Seen Rx],0) as [OSTOCAL Seen Rx],
        isnull([SOLBION Seen Rx],0) as [SOLBION Seen Rx],
        isnull([XINC Seen Rx],0) as [XINC Seen Rx],
        isnull([Maxfer Seen Rx],0) as [Maxfer Seen Rx],
		isnull([Zeefol Seen Rx],0) as [Zeefol Seen Rx],
		isnull([Danamet Seen Rx],0) as [Danamet Seen Rx],
		isnull([Ethinor Seen Rx],0) as [Ethinor Seen Rx],
		isnull([Lenor Seen Rx],0) as [Lenor Seen Rx],
		isnull([Zatral Seen Rx],0) as [Zatral Seen Rx],
		isnull([Urokit Seen Rx],0) as [Urokit Seen Rx],
		isnull([Valenty Seen Rx],0) as [Valenty Seen Rx]
              
        from v_LastDay_SeenRxAB ) as T1
        -- where [FFTR] not in ('IN', 'INST')
        order by [FFTR] asc

     """, conn.m_reporting)
    new_seen_rx_data.to_excel("./Data/SeenRx/Seen_Rx_Data.xlsx", index=False)
    # data = pd.read_excel('./Data/SeenRx/Seen_Rx_Data.xlsx')

    # CBU = data[data['FFTR'].str.contains('CBU')]
    # CBU.to_excel('./Data/SeenRx/SeenRx_CBU.xlsx', index=False)
    #
    # CCF = data[data['FFTR'].str.contains('CCF')]
    # CCF.to_excel('./Data/SeenRx/SeenRx_CCF.xlsx', index=False)
    #
    # CCX = data[data['FFTR'].str.contains('CCX')]
    # CCX.to_excel('./Data/SeenRx/SeenRx_CCX.xlsx', index=False)
    #
    # CNH = data[data['FFTR'].str.contains('CNH')]
    # CNH.to_excel('./Data/SeenRx/SeenRx_CNH.xlsx', index=False)
    #
    # CKJ = data[data['FFTR'].str.contains('CKJ')]
    # CKJ.to_excel('./Data/SeenRx/SeenRx_CKJ.xlsx', index=False)
    #
    # CMT = data[data['FFTR'].str.contains('CMT')]
    # CMT.to_excel('./Data/SeenRx/SeenRx_CMT.xlsx', index=False)
    #
    # CRB = data[data['FFTR'].str.contains('CRB')]
    # CRB.to_excel('./Data/SeenRx/SeenRx_CRB.xlsx', index=False)
    #
    # CRP = data[data['FFTR'].str.contains('CRP')]
    # CRP.to_excel('./Data/SeenRx/SeenRx_CRP.xlsx', index=False)
    #
    # CSB = data[data['FFTR'].str.contains('CSB')]
    # CSB.to_excel('./Data/SeenRx/SeenRx_CSB.xlsx', index=False)
    print('2. Seen Rx Data Saved')


def doctor_call_data():
    doctor_call_data = pd.read_sql_query(""" 
        select * from
        (Select
        left([FF ID],2) as [FFTR],
        isnull(sum([OSTOCAL Call]),0)  as [OSTOCAL],
        isnull(sum([SOLBION Call]),0) as [SOLBION],
        isnull(sum([XINC Call]),0) as [XINC],
        isnull(sum([MAXFER Call]),0) as [MAXFER Call],
        isnull(sum([ZEEFOL Call]),0) as [ZEEFOL Call],
        isnull(sum([DANAMET Call]),0) as [DANAMET Call],
        isnull(sum([ETHINOR Call]),0) as [ETHINOR Call],
        isnull(sum([LENOR Call]),0) as [LENOR Call],
        isnull(sum([ZATRAL Call]),0) as [ZATRAL Call],
        isnull(sum([UROKIT Call]),0) as [UROKIT Call],
        isnull(sum([VALENTY Call]),0) as [VALENTY Call]
                               
        from [V_LastDayDoctorCallAB]
        --where [FF ID]  = left([FF ID],3)+'0'
        group by left([FF ID],2)
        union all
                                
        Select [FF ID], 
            ISNULL( [OSTOCAL Call], 0) AS [OSTOCAL Call],
            ISNULL( [SOLBION Call],0) AS [SOLBION Call],
            ISNULL([XINC Call],0) AS [XINC Call],
            ISNULL([MAXFER Call], 0) AS [MAXFER Call],
            ISNULL( [ZEEFOL Call], 0) AS [ZEEFOL Call],
            ISNULL( [DANAMET Call], 0) AS [DANAMET Call], 
            ISNULL([ETHINOR Call],0) AS [ETHINOR Call], 
            ISNULL([LENOR Call],0) AS [LENOR Call],
            ISNULL( [ZATRAL Call],0) AS  [ZATRAL Call], 
            ISNULL( [UROKIT Call],0) AS  [UROKIT Call], 
            ISNULL( [VALENTY Call], 0) AS  [VALENTY Call]   
        from [V_LastDayDoctorCallAB]                
        ) as T1
        -- where FFTR not in ('IN')
        order by [FFTR] asc
         """, conn.m_reporting)
    doctor_call_data.to_excel("./Data/Call/doctor_call_data.xlsx", index=False)

    # data = pd.read_excel('./Data/Call/doctor_call_data.xlsx')

    # CBU = data[data['FFTR'].str.contains('CBU')]
    # CBU.to_excel('./Data/Call/Call_CBU.xlsx', index=False)
    #
    # CCF = data[data['FFTR'].str.contains('CCF')]
    # CCF.to_excel('./Data/Call/Call_CCF.xlsx', index=False)
    #
    # CCX = data[data['FFTR'].str.contains('CCX')]
    # CCX.to_excel('./Data/Call/Call_CCX.xlsx', index=False)
    #
    # CNH = data[data['FFTR'].str.contains('CNH')]
    # CNH.to_excel('./Data/Call/Call_CNH.xlsx', index=False)
    #
    # CKJ = data[data['FFTR'].str.contains('CKJ')]
    # CKJ.to_excel('./Data/Call/Call_CKJ.xlsx', index=False)
    #
    # CMT = data[data['FFTR'].str.contains('CMT')]
    # CMT.to_excel('./Data/Call/Call_CMT.xlsx', index=False)
    #
    # CRB = data[data['FFTR'].str.contains('CRB')]
    # CRB.to_excel('./Data/Call/Call_CRB.xlsx', index=False)
    #
    # CRP = data[data['FFTR'].str.contains('CRP')]
    # CRP.to_excel('./Data/Call/Call_CRP.xlsx', index=False)
    #
    # CSB = data[data['FFTR'].str.contains('CSB')]
    # CSB.to_excel('./Data/Call/Call_CSB.xlsx', index=False)

    print('3. Doctor Call Data Saved')
