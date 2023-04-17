import imaplib
import pandas as pd 
import email
import io
import zipfile
from zipfile import BadZipfile

# Email login credentials
EMAIL_ACCOUNT = "himanshumanral2003@gmail.com"
PASSWORD = "qvicfgwfprwbnlrd"

# IMAP server settings for Gmail
SERVER = "imap.gmail.com"
PORT = 993

# Connect to the server and login with 30 second timeout
mail = imaplib.IMAP4_SSL(SERVER, PORT, timeout=30)
mail.login(EMAIL_ACCOUNT, PASSWORD)

mail.select('inbox')

# Search for emails based on their internal date
result, data = mail.search(None, 'SINCE "01-Apr-2023"')

# Retrieve the latest email based on its UID
latest_email_uid = data[0].split()[-1]
result, data = mail.fetch(latest_email_uid, "(RFC822)")

# Parse the email data using the email library
raw_email = data[0][1]
email_message = email.message_from_bytes(raw_email)

# Find and extract the Excel attachment from the email
attachment_data = None
for part in email_message.walk():
    content_type = part.get_content_type()
    if content_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        attachment_data = part.get_payload(decode=True)
        break
    elif content_type == 'multipart/mixed':
        # Check for attachments in the sub-parts of a multipart email
        for subpart in part.walk():
            subpart_type = subpart.get_content_type()
            if subpart_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                attachment_data = subpart.get_payload(decode=True)
                break

# Read the Excel attachment data as a pandas dataframe
if attachment_data is not None:
    df = pd.read_excel(io.BytesIO(attachment_data))
    print(df.head())
else:
    print("No Excel attachment found in the latest email.")

mail.close()
mail.logout()       
# Do something with the mailbox, for example select the Inbox folder
# mail.select("INBOX")

# # Close the mailbox and logout
# mail.close()
 
# from_mail = 'FROM "himanshumanral2003@gmail.com"'

# mail.select('inbox')

# # Search for emails based on their internal date
# result, data = mail.search(None,from_mail)

# # Retrieve the latest email based on its UID
# latest_email_uid = data[0].split()[-1]
# result, data = mail.fetch(latest_email_uid, "(RFC822)")

# # Parse the email data using the email library
# raw_email = data[0][1]
# email_message = email.message_from_bytes(raw_email)

# # Find and extract the zip attachment from the email
# zipfile_data = None
# for part in email_message.walk():
#     if part.get_content_type() == 'application/zip':
#         try:
#             zipfile_data = io.BytesIO(part.get_payload(decode=True))
#             with zipfile.ZipFile(zipfile_data) as z:
#                 file_list = z.namelist()
#         except BadZipfile:
#             print("ERROR: File is not a valid zip file.")
#             break

# if zipfile_data:
#     # Check that the zip file contains an Excel file
#     excel_file = None
#     for file_name in file_list:
#         if file_name.endswith(".xlsx"):
#             excel_file = file_name
#             print("hello")

#     if excel_file:
#         # Read the Excel attachment data as a pandas dataframe
#         with zipfile.ZipFile(zipfile_data) as z:
#             with z.open(excel_file) as f:
#                 df = pd.read_excel(f)
#         print(df.head())
#     else:
#         print("No Excel file found in the zip attachment.")
# else:
#     print("No zip attachment found in the latest email.")
