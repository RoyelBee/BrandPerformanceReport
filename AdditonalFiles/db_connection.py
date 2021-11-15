
import pyodbc as db

m_reporting = db.connect('DRIVER={SQL Server};'
                         'SERVER=10.168.2.247;'
                         'DATABASE=DCR_MREPORTING;'
                         'UID=sa;PWD=erp;')

azure = db.connect('DRIVER={SQL Server};'
                   'SERVER=137.116.139.217;'
                   'DATABASE=ARCHIVESKF;'
                   'UID=sa;PWD=erp@123;')