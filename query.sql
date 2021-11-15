
-------------- Sales Achievement and Trend percentage query ------------------------------------
----------------------------------------------------------------------------------------------
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
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'LIGAZID' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [LIGAZID SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'LIGAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [LIGAZID TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'EMAZID' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [EMAZID SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'EMAZID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [EMAZID TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'LIPICON' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [LIPICON SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'LIPICON' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [LIPICON TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AGLIP' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [AGLIP SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AGLIP' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [AGLIP TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'CIFIBET' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [CIFIBET SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'CIFIBET' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [CIFIBET TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AMLEVO' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [AMLEVO SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AMLEVO' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [AMLEVO TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'CARDOBIS' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [CARDOBIS SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'CARDOBIS' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [CARDOBIS TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Rivarox' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [RIVAROX SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Rivarox' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [RIVAROX TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'NOCLOG' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [NOCLOG SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'NOCLOG' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [NOCLOG TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'BEMPID' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [BEMPID SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'BEMPID' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [BEMPID TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AROTIDE' THEN (SalesAmount/TargetAmount)*100 END),0)) AS [AROTIDE SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'AROTIDE' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [AROTIDE TREND%],

Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'FOBUNID' THEN (SalesAmount/TargetAmount) *100 END),0)) AS [FOBUNID SALES ACHIV%],
Convert(decimal(18,2), isnull(Sum(Case when FFTarget.Brand = 'Fobunid' THEN (SalesAmount/@TotalDaysGone*@TotalDaysInMonth)/TargetAmount *100 END),0)) AS [FOBUNID TREND%]

from
(
Select 'RSM' AS FF, RSMTR as FFTR,BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target
group by RSMTR,BRAND
union all
Select 'FM' AS FF, FMTR as FFTR,BRAND,SUM(Amount) AS TargetAmount  from V_FF_Brand_Target
group by FMTR,BRAND
union all
Select 'MSO' AS FF, MSOTR as FFTR,BRAND, SUM(Amount) AS TargetAmount  from V_FF_Brand_Target
group by MSOTR,BRAND
) as FFTarget
left join
(
Select 'RSM' AS FF, left(MSOTR,3)  as FFTR,BRAND, SUM(extinvmisc) AS SalesAmount  from V_FF_Brand_Sales
where TRANSDATE <= @LastDay
group by left(MSOTR,3) ,BRAND
union all
Select 'FM' AS FF, left(MSOTR,4)+'0'  as FFTR,BRAND,SUM(extinvmisc) AS SalesAmount   from V_FF_Brand_Sales
where TRANSDATE <= @LastDay
group by left(MSOTR,4)+'0' ,BRAND
union all
Select 'MSO' AS FF, MSOTR as FFTR,BRAND, SUM(extinvmisc) AS SalesAmount   from V_FF_Brand_Sales
where TRANSDATE <= @LastDay
group by MSOTR,BRAND
) as FFSales
ON (FFTarget.FFTR=FFSales.FFTR) AND (FFTarget.BRAND=FFSales.BRAND)
group by FFTarget.FFTR
) as T1
order by FFTR asc



------------------------------------------------------------------------------------------------------
----------------------------------- Last Day Seen Rx -------------------------------------------------
------------------------------------------------------------------------------------------------------
select * from
(Select
left([FF ID],3) as [FFTR],
isnull(sum([LIGAZID Seen Rx]),0)  as [LIGAZID],
isnull(sum([EMAZID Seen Rx]),0) as [EMAZID],
isnull(sum([LIPICON Seen Rx]),0) as [LIPICON],
isnull(sum([AGLIP Seen Rx]),0) as [AGLIP],
isnull(sum([CIFIBET Seen Rx]),0) as [CIFIBET],
isnull(sum([AMLEVO Seen Rx]),0) as [AMLEVO],
isnull(sum([CARDOBIS Seen Rx]),0) as [CARDOBIS],
isnull(sum([RIVAROX Seen Rx]),0) as [RIVAROX],
isnull(sum([NOCLOG Seen Rx]),0) as [NOCLOG],
isnull(sum([BEMPID Seen Rx]),0) as [BEMPID],
isnull(sum([AROTIDE Seen Rx]),0) as [AROTIDE],
isnull(sum([FOBUNID Seen Rx]),0) as [FOBUNID]

from v_LastDay_SeenRx
where [FF ID]  = left([FF ID],4)+'0' --and left([FF ID],3) like '%CBU%'
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
ISNULL([NOCLOG Seen Rx], 0) as [NOCLOG Seen Rx],
ISNULL([BEMPID Seen Rx], 0) as [BEMPID Seen Rx],
ISNULL([AROTIDE Seen Rx], 0) as [AROTIDE Seen Rx],
ISNULL([FOBUNID Seen Rx], 0) as [FOBUNID Seen Rx]
from v_LastDay_SeenRx --where  left([FF ID],3) like '%CBU%'

) as T1
order by [FFTR] asc

------------------------------------------------------------------------------
---------------------------- Doctor Call -------------------------------------
------------------------------------------------------------------------------

select * from
(Select
left([FF ID],3) as [FFTR],
isnull(sum([LIGAZID Call]),0)  as [LIGAZID],
isnull(sum([EMAZID Call]),0) as [EMAZID],
isnull(sum([LIPICON Call]),0) as [LIPICON],
isnull(sum([AGLIP Call]),0) as [AGLIP],
isnull(sum([CIFIBET Call]),0) as [CIFIBET],
isnull(sum([AMLEVO Call]),0) as [AMLEVO],
isnull(sum([CARDOBIS Call]),0) as [CARDOBIS],
isnull(sum([RIVAROX Call]),0) as [RIVAROX],
isnull(sum([NOCLOG Call]),0) as [NOCLOG],
isnull(sum([BEMPID Call]),0) as [BEMPID],
isnull(sum([AROTIDE Call]),0) as [AROTIDE],
isnull(sum([FOBUNID Call]),0) as [FOBUNID]

from [V_LastDayDoctorCall]
where [FF ID]  = left([FF ID],4)+'0'
group by left([FF ID],3)
union all

Select [FF ID],
isnull([LIGAZID Call],0) as [LIGAZID Call],
isnull([EMAZID Call],0) as [EMAZID Call],
ISNULL([LIPICON Call], 0) as [LIPICON Call],
ISNULL([AGLIP Call], 0) as [AGLIP Call],
ISNULL([CIFIBET Call], 0) as [CIFIBET Call],
ISNULL([AMLEVO Call], 0) as [AMLEVO Call],
ISNULL([CARDOBIS Call], 0) as [CARDOBIS Call],
ISNULL([RIVAROX Call], 0) as [RIVAROX Call],
ISNULL([NOCLOG Call], 0) as [NOCLOG Call],
ISNULL([BEMPID Call], 0) as [BEMPID Call],
ISNULL([AROTIDE Call], 0) as [AROTIDE Call],
ISNULL([FOBUNID Call], 0) as [FOBUNID Call]
from [V_LastDayDoctorCall]

) as T1
order by [FFTR] asc

