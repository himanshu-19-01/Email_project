
import email
import imaplib
import pandas as pd
import io
# Email login credancials
EMAIL_ACCOUNT = "himanshumanral2003@gmail.com"
PASSWORD = "qvicfgwfprwbnlrd"
server_1 = "imap.gmail.com"


mail = imaplib.IMAP4(server_1,port=993,timeout=60)
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
for part in email_message.walk():
    if part.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
        attachment_data = part.get_payload(decode=True)
        break

# Read the Excel attachment data as a pandas dataframe
if attachment_data:
    df = pd.read_excel(io.BytesIO(attachment_data))
    print(df.head())
else:
    print("No Excel attachment found in the latest email.")
mail.close()
mail.logout()    