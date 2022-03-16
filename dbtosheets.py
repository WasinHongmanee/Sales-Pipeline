import gspread
from pprint import pprint
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import pyodbc


conx_string = 'DRIVER={SQL Server Native Client 11.0}; SERVER=localhost; Database=FoursoftDb; TRUSTED_CONNECTION=yes'
allsales = """
        select FaceAmount as Sales,CreatedWhen as Date from dbo.OrderPayment
        where CreatedWhen DATEADD(day, -90, GETDATE()) AND GETDATE()
        order by CreatedWhen DESC
        """

deliverysales = """
        select FaceAmount as Sales,OrderPayment.CreatedWhen as Date from dbo.OrderPayment
        join dbo.OrderList on
        OrderList.id = OrderPayment.OrderId
        where OrderPayment.CreatedWhen BETWEEN DATEADD(day, -90, GETDATE()) AND GETDATE()
        and OrderType = 62
        """

onlinepickupsales = """
        select FaceAmount as Sales,OrderPayment.CreatedWhen as Date from dbo.OrderPayment
        join dbo.OrderList on
        OrderList.id = OrderPayment.OrderId
        where OrderPayment.CreatedWhen BETWEEN DATEADD(day, -90, GETDATE()) AND GETDATE()
        and OrderType = 63
        and PaidMethodName = 'Online Payment'
"""
conx = pyodbc.connect(conx_string) #connect to the database
allsalesdf = pd.read_sql(sql = allsales, con = conx)
deliverysalesdf = pd.read_sql(sql = deliverysales, con = conx)
onlinepickupsalesdf = pd.read_sql(sql = onlinepickupsales, con = conx)

conx.close() #close the connection


scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credsforpeeps.json', scope)
client = gspread.authorize(creds)

sheet1 = client.open('dashboardo').worksheet('Sheet1')
sheet2 = client.open('dashboardo').worksheet('Sheet2')
sheet3 = client.open('dashboardo').worksheet('Sheet3')
sheet1.clear() #clear the google sheet
sheet2.clear()
sheet3.clear()

allsalesdf['Date'] = allsalesdf['Date'].astype(str)
deliverysalesdf['Date'] = deliverysalesdf['Date'].astype(str)
onlinepickupsalesdf['Date'] = onlinepickupsalesdf['Date'].astype(str)

sheet1.update([allsalesdf.columns.values.tolist()] + allsalesdf.values.tolist())
sheet2.update([deliverysalesdf.columns.values.tolist()] + deliverysalesdf.values.tolist())
sheet3.update([onlinepickupsalesdf.columns.values.tolist()] + onlinepickupsalesdf.values.tolist())
