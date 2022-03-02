import gspread
import sqlite3
from pprint import pprint
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd



scope = ["https://spreadsheets.google.com/feeds",
         'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('credsforpeeps.json', scope)
client = gspread.authorize(creds)
sheet1 = client.open('dashboardo').worksheet('Sheet1')
sheet2 = client.open('dashboardo').worksheet('Sheet2')

#all_records = sheet1.get_all_records()
#pprint(all_records)

sheet1.clear() #clear the google sheet

df2 = pd.read_csv('/Users/wasin/Downloads/sheetsforpeeps/7days.csv')
df2 = df2[['CreatedWhen','id','OrderId','DueAmount']]
df = pd.read_csv('/Users/wasin/Downloads/sheetsforpeeps/results.csv')
print(df.columns)
df = df[['CreatedWhen','id','OrderId','DueAmount']]
print(df.head())
sheet1.update([df.columns.values.tolist()] + df.values.tolist())
sheet2.update([df2.columns.values.tolist()] + df2.values.tolist())


'''
try:
    connection = sqlite3.connect('FoursoftDb.bak')
    print('connected to DB')
except Exception as error:
    print('Error code:', str(error))
    
connection.close()
'''