import smtplib
import pyodbc
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import io

# connect to the database and fetch the data
server = 'DESKTOP-SAAQPQ1'
database = 'Dashbord'
username = 'sa'
password = '914913'
conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')
query="SELECT * FROM Users"

# convert the data into a Pandas DataFrame
df = pd.read_sql(query, conn)
 
 

# create the Excel file in memory
excel_file = io.BytesIO()
writer = pd.ExcelWriter(excel_file, engine='openpyxl')
df.to_excel(writer, index=False)
writer.close()
excel_file.seek(0)
# create the email message
msg = MIMEMultipart()
msg['From'] = "himanshumanral2003@gmail.com"
# msg['To'] = "deepmanral2004@gmail.com, nareshsinghankit@gmail.com"
msg['To']="himanshumanral2003@gmail.com"
msg['Subject'] = "Database data attachment"

# attach the Excel file to the email
attachment = MIMEBase("application", "octet-stream")
attachment.set_payload(excel_file.read())
encoders.encode_base64(attachment)
attachment.add_header("Content-Disposition", "attachment", filename="my_data.xlsx")
msg.attach(attachment)

# add the email body
body = "This is a system generated email."
msg.attach(MIMEText(body, 'plain'))

# send the email
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_username = "himanshumanral2003@gmail.com"
smtp_password = "qvicfgwfprwbnlrd"

with smtplib.SMTP(smtp_server, smtp_port) as smtp:
    smtp.starttls()
    smtp.login(smtp_username, smtp_password)
    smtp.sendmail(smtp_username, msg['To'].split(','), msg.as_string())

print("Email sent successfully!")
