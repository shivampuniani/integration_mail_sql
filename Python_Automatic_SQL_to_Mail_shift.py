import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import pyodbc

import configparser


def get_db_config():
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    db_config = {
        'server': config.get('database', 'server'),
        'database': config.get('database', 'database'),
        'username': config.get('database', 'username'),
        'password': config.get('database', 'password'),
        'smtp_server': config.get('smtp', 'smtp_server'),
        'port': config.get('smtp', 'port'),
        'from': config.get('recipients', 'from'),
        'toRecipients': config.get('recipients', 'toRecipients'),
        'ccRecipients': config.get('recipients', 'ccRecipients')
    }
    return db_config



db_config = get_db_config()
con = pyodbc.connect(
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={db_config["server"]};'
        f'DATABASE={db_config["database"]};'
        f'UID={db_config["username"]};'
        f'PWD={db_config["password"]}'
)

smtp_server = db_config["smtp_server"]
port = db_config["port"]

toRecipients = db_config["toRecipients"]
ccRecipients = db_config["ccRecipients"]

FROM = db_config["from"]
#TO = "shivam@tridentindia.com"

def getDataFromSql(sqlString, outputSqlData, outputListData, sqlStringName):
    outputSqlData = cur.execute(sqlString)
    for row in outputSqlData:
        print(sqlStringName)
        for i in row:
            #print(i)
            outputListData.append(str(i))

today = datetime.now()
yesterday = today - timedelta(days=1)
now = str(today.strftime("%H:%M:%S"))
if (now > "00:00:00" and now < "06:59:59"):
    Shift = "B"
    Start_time = str(yesterday.strftime("%Y-%m-%d 14:58:00"))
    End_time = str(yesterday.strftime("%Y-%m-%d 23:00:00"))
elif (now >= "23:00:00" and now < "23:59:59"):
    Shift = "B"
    Start_time = str(today.strftime("%Y-%m-%d 14:58:00"))
    End_time = str(today.strftime("%Y-%m-%d 23:00:00"))
elif (now > "07:00:00" and now < "14:59:59"):
    Shift = "C"
    Start_time = str(yesterday.strftime("%Y-%m-%d 22:58:00"))
    End_time = str(today.strftime("%Y-%m-%d 07:00:00"))
elif (now > "15:00:00" and now < "22:59:59"):
    Shift = "A"
    Start_time = str(today.strftime("%Y-%m-%d 06:58:00"))
    End_time = str(today.strftime("%Y-%m-%d 15:00:00"))

now2 = (End_time[:11])



print(Start_time+'_'+End_time + " " +now2)

cur = con.cursor()

sql_Mail_X_All_Value = """ SELECT TOP(10)
       [Mail_X_Field1_Actual_Value]
      ,[Mail_X_Field2_Actual_Value]
      ,[Mail_X_Field3_Actual_Value]
      ,[Mail_X_Field4_Actual_Value]
      ,[Mail_X_Field5_Actual_Value]
      ,[Mail_X_Field6_Actual_Value]
      ,[Mail_X_Field7_Actual_Value]
      ,[Mail_X_Field8_Actual_Value]
      ,[Mail_X_Field9_Actual_Value]
      ,[Mail_X_Field10_Actual_Value]
      ,[Mail_X_Field11_Actual_Value]
      ,[Mail_X_Field12_Actual_Value]
      , Timestamp
  FROM [dbo].[Mail_X]  order by Timestamp asc"""


sql_Mail_Y_All_Value = """  SELECT TOP(10)
       [Mail_Y_Field1_Actual_Value]
      ,[Mail_Y_Field2_Actual_Value]
      ,[Mail_Y_Field3_Actual_Value]
      ,[Mail_Y_Field4_Actual_Value]
      ,[Mail_Y_Field5_Actual_Value]
      ,[Mail_Y_Field6_Actual_Value]
      ,[Mail_Y_Field7_Actual_Value]
      ,[Mail_Y_Field8_Actual_Value]
      ,[Mail_Y_Field9_Actual_Value]
      ,[Mail_Y_Field10_Actual_Value]
      ,[Mail_Y_Field11_Actual_Value]
      ,[Mail_Y_Field12_Actual_Value]
      , Timestamp
  FROM [dbo].[Mail_Y] order by Timestamp asc"""


sql_Mail_Y_Target_Value = """

SELECT TOP (1)
        [Mail_Y_Field1_Target_Value]
      ,[Mail_Y_Field2_Target_Value]
      ,[Mail_Y_Field3_Target_Value]
      ,[Mail_Y_Field4_Target_Value]
      ,[Mail_Y_Field5_Target_Value]
      ,[Mail_Y_Field6_Target_Value]
      ,[Mail_Y_Field7_Target_Value]
      ,[Mail_Y_Field8_Target_Value]
      ,[Mail_Y_Field9_Target_Value]
      ,[Mail_Y_Field10_Target_Value]
      ,[Mail_Y_Field11_Target_Value]
      ,[Mail_Y_Field12_Target_Value]
  FROM [dbo].[Mail_Y]  order by Timestamp asc"""

sql_Mail_X_Target_Value = """
SELECT TOP (1)
       [Mail_X_Field1_Target_Value]
      ,[Mail_X_Field2_Target_Value]
      ,[Mail_X_Field3_Target_Value]
      ,[Mail_X_Field4_Target_Value]
      ,[Mail_X_Field5_Target_Value]
      ,[Mail_X_Field6_Target_Value]
      ,[Mail_X_Field7_Target_Value]
      ,[Mail_X_Field8_Target_Value]
      ,[Mail_X_Field9_Target_Value]
      ,[Mail_X_Field10_Target_Value]
      ,[Mail_X_Field11_Target_Value]
      ,[Mail_X_Field12_Target_Value]
  FROM [dbo].[Mail_X]  order by Timestamp asc"""


Mail_X_All_row_sql_data = cur.execute(sql_Mail_X_All_Value)
Mail_X_Target_row_sql_data = cur.execute(sql_Mail_X_Target_Value)

Mail_Y_All_row_sql_data = cur.execute(sql_Mail_Y_All_Value)
Mail_Y_Target_row_sql_data = cur.execute(sql_Mail_Y_Target_Value)


Mail_X_All_row = []
Mail_X_Target_row = []

Mail_Y_All_row = []
Mail_Y_Target_row = []


getDataFromSql(sql_Mail_X_All_Value,Mail_X_All_row_sql_data,Mail_X_All_row,"Mail_X_All_row_sql_data")
getDataFromSql(sql_Mail_X_Target_Value,Mail_X_Target_row_sql_data,Mail_X_Target_row,"Mail_X_Target_row_sql_data")

getDataFromSql(sql_Mail_Y_All_Value,Mail_Y_All_row_sql_data,Mail_Y_All_row,"Mail_Y_All_row_sql_data")
getDataFromSql(sql_Mail_Y_Target_Value,Mail_Y_Target_row_sql_data,Mail_Y_Target_row,"Mail_Y_Target_row_sql_data")

cur.close()

Mail_Y_even_fields = 1
Mail_Y_odd_fields = 2.5
Mail_X_even_fields = 1
Mail_X_odd_fields = 2.5

def checkTolerance(actual,tol, setValue):
    if (float(setValue) + tol < float(actual)) or (float(setValue) - tol > float(actual)) :
        if(tol==2.5):
            return "bgcolor= '69cafe'"
        else:
            return "bgcolor= 'face69'"
    else:
        return ""

print(checkTolerance(Mail_X_Target_row[8],Mail_X_odd_fields,Mail_X_All_row[21]))
print(checkTolerance(Mail_X_Target_row[8],Mail_X_odd_fields,Mail_X_All_row[34]))

SUBJECT = "Unit Report of "+  now2 +", " + Shift + "  shift!"
TEXT = "Dear Team, Please find attached the "

html = """

<html>
<body>

<p>Dear Team,

<p>Please find below the summary of Unit report of """ +  now2 +' for the '+ Shift + """  shift.


,




<table width: 500px border='1' cellpadding='2' style=' border:1px solid black; border-collapse:collapse'>

<thead><tr><th colspan='15'; bgcolor = #008775><h2>
<div style='text-align: center; color:#ffffff'><span style='font-family: Calibri'>Unit Report</span></div></h2>
</th></tr></thead>


<tbody>
<tr bgcolor = #fdb612>
<td style='width: 30px;  text-align: center; color:#000000'><span style='font-family: Calibri'>Sr No. </span></td>
<td style='width: 100px; text-align: center; color:#000000'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center; color:#000000'><span style='font-family: Calibri'>Mail_X/Mail_Y</span></td>
<td style='width: 100px; text-align: center; color:#000000'><span style='font-family: Calibri'>Even Fields/Odd Fields</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>Target Value</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>Tolerance Value</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[12] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[25] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[38] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[51] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[64] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[77] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[90] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[103] + """</span></td>
<td style='width: 80px;  text-align: center; color:#000000'><span style='font-family: Calibri'>""" + Mail_X_All_row[116] + """</span></td>
</tr>
<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>1</span></td>
<td style='width: 100px; text-align: center'; rowspan = '4'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_X</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[0] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance( Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[0])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[0] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[13])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[13] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[26])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[26] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[39])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[39] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[52])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[52] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[65])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[65] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[78])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[78] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[91])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[91] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[0], Mail_X_odd_fields, Mail_X_All_row[104])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[104] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>2</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[1] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[1])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[1] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[14])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[14] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[27])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[27] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[40])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[40] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[53])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[53] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[66])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[66] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[79])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[79] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[92])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[92] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[1], Mail_X_even_fields, Mail_X_All_row[105])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[105] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>3</span></td>

<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_Y</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[0] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[0])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[0] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[13])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[13] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[26])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[26] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[39])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[39] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[52])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[52] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[65])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[65] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[78])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[78] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[91])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[91] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[0], Mail_Y_odd_fields, Mail_Y_All_row[104])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[104] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>4</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[1] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[1])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[1] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[14])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[14] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[27])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[27] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[40])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[40] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[53])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[53] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[66])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[66] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[79])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[79] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[92])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[92] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[1], Mail_Y_even_fields, Mail_Y_All_row[105])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[105] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>5</span></td>
<td style='width: 100px; text-align: center'; rowspan = '4'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_X</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[2] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[2])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[2] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[15])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[15] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[28])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[28] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[41])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[41] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[54])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[54] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[67])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[67] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[80])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[80] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[93])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[93] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[2], Mail_X_odd_fields, Mail_X_All_row[106])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[106] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>6</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[3] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[3])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[3] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[16])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[16] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[29])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[29] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[42])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[42] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[55])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[55] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[68])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[68] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[81])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[81] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[94])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[94] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[3], Mail_X_even_fields, Mail_X_All_row[107])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[107] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>7</span></td>

<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_Y</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[2] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[2])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[2] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[15])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[15] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[28])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[28] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[41])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[41] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[54])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[54] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[67])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[67] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[80])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[80] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[93])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[93] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[2], Mail_Y_odd_fields, Mail_Y_All_row[106])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[106] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>8</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[3] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[3])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[3] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[16])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[16] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[29])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[29] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[42])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[42] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[55])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[55] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[68])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[68] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[81])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[81] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[94])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[94] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[3], Mail_Y_even_fields, Mail_Y_All_row[107])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[107] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>9</span></td>
<td style='width: 100px; text-align: center'; rowspan = '4'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_X</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[4] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[4])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[4] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[17])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[17] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[30])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[30] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[43])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[43] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[56])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[56] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[69])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[69] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[82])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[82] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[95])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[95] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[4], Mail_X_odd_fields, Mail_X_All_row[108])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[108] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>10</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[5] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[5])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[5] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[18])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[18] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[31])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[31] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[44])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[44] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[57])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[57] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[70])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[70] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[83])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[83] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[96])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[96] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[5], Mail_X_even_fields, Mail_X_All_row[109])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[109] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>11</span></td>

<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_Y</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[4] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[4])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[4] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[17])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[17] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[30])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[30] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[43])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[43] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[56])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[56] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[69])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[69] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[82])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[82] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[95])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[95] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[4], Mail_Y_odd_fields, Mail_Y_All_row[108])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[108] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>12</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[5] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[5])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[5] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[18])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[18] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[31])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[31] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[44])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[44] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[57])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[57] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[70])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[70] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[83])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[83] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[96])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[96] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[5], Mail_Y_even_fields, Mail_Y_All_row[109])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[109] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>13</span></td>
<td style='width: 100px; text-align: center'; rowspan = '4'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_X</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[6] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[6])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[6] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[19])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[19] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[32])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[32] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[45])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[45] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[58])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[58] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[71])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[71] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[84])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[84] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[97])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[97] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[6], Mail_X_odd_fields, Mail_X_All_row[110])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[110] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>14</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[7] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[7])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[7] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[20])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[20] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[33])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[33] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[46])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[46] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[59])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[59] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[72])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[72] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[85])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[85] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[98])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[98] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[7], Mail_X_even_fields, Mail_X_All_row[111])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[111] + """</span></td>
</tr>

"<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>15</span></td>"

<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_Y</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[6] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[6])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[6] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[19])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[19] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[32])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[32] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[45])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[45] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[58])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[58] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[71])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[71] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[84])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[84] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[97])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[97] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[6], Mail_Y_odd_fields, Mail_Y_All_row[110])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[110] + """</span></td>
</tr>

"<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>16</span></td>"


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[7] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[7])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[7] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[20])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[20] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[33])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[33] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[46])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[46] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[59])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[59] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[72])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[72] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[85])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[85] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[98])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[98] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[7], Mail_Y_even_fields, Mail_Y_All_row[111])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[111] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>17</span></td>
<td style='width: 100px; text-align: center'; rowspan = '4'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_X</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[8] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[8])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[8] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[21])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[21] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[34])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[34] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[47])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[47] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[60])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[60] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[73])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[73] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[86])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[86] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[99])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[99] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[8], Mail_X_odd_fields, Mail_X_All_row[112])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[112] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>18</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[9] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[9])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[9] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[22])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[22] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[35])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[35] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[48])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[48] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[61])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[61] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[74])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[74] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[87])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[87] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[100])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[100] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[9], Mail_X_even_fields, Mail_X_All_row[113])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[113] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>19</span></td>

<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_Y</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[8] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[8])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[8] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[21])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[21] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[34])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[34] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[47])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[47] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[60])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[60] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[73])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[73] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[86])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[86] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[99])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[99] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[8], Mail_Y_odd_fields, Mail_Y_All_row[112])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[112] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>20</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[9] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[9])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[9] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[22])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[22] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[35])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[35] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[48])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[48] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[61])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[61] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[74])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[74] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[87])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[87] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[100])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[100] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[9], Mail_Y_even_fields, Mail_Y_All_row[113])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[113] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>21</span></td>
<td style='width: 100px; text-align: center'; rowspan = '4'><span style='font-family: Calibri'>Fields</span></td>
<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_X</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[10] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[10])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[10] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[23])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[23] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[36])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[36] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[49])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[49] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[62])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[62] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[75])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[75] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[88])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[88] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[101])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[101] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[10], Mail_X_odd_fields, Mail_X_All_row[114])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[114] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>22</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_X_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_Target_row[11] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[11])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[11] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[24])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[24] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[37])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[37] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[50])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[50] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[63])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[63] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[76])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[76] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[89])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[89] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[102])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[102] + """</span></td>
<td """ + (checkTolerance(Mail_X_Target_row[11], Mail_X_even_fields, Mail_X_All_row[115])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_X_All_row[115] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>23</span></td>

<td style='width: 70px;  text-align: center'; rowspan = '2'><span style='font-family: Calibri'>Mail_Y</span></td>
<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_odd_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[10] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 2.5</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[10])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[10] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[23])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[23] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[36])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[36] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[49])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[49] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[62])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[62] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[75])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[75] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[88])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[88] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[101])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[101] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[10], Mail_Y_odd_fields, Mail_Y_All_row[114])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[114] + """</span></td>
</tr>

<tr>
<td style='width: 30px;  text-align: center'><span style='font-family: Calibri'>24</span></td>


<td style='width: 100px; text-align: center'><span style='font-family: Calibri'>Mail_Y_even_fields</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_Target_row[11] + """</span></td>
<td style='width: 80px;  text-align: center'><span style='font-family: Calibri'>± 1.0</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[11])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[11] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[24])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[24] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[37])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[37] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[50])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[50] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[63])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[63] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[76])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[76] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[89])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[89] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[102])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[102] + """</span></td>
<td """ + (checkTolerance(Mail_Y_Target_row[11], Mail_Y_even_fields, Mail_Y_All_row[115])) + """ style='width: 80px;  text-align: center'><span style='font-family: Calibri'>""" + Mail_Y_All_row[115] + """</span></td>
</tr>

</tbody>

</table>
<p>Kind regards,<br>
</body></html>

"""


part1 = MIMEText(TEXT, 'plain')
part2 = MIMEText(html, 'html')

msg = MIMEMultipart('alternative')
msg['Subject'] = SUBJECT
msg['From'] = FROM
#msg['To'] = ", ".join(toRecipients)
#msg['Cc'] = ", ".join(ccRecipients)
msg['To'] = (toRecipients)
msg['Cc'] = (ccRecipients)
msg.attach(part1)
msg.attach(part2)

context = ssl.create_default_context()
server = smtplib.SMTP(smtp_server,port)
TO = toRecipients + ccRecipients
server.sendmail(FROM, TO, msg.as_string())
server.quit()
