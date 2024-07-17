import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

# Assuming the files are in the same directory as the script
excel_files = [os.path.abspath("file1.xlsx"), os.path.abspath("file2.xlsx"), os.path.abspath("file3.xlsx")]


# Load Excel spreadsheets and combine into one
excel_files = ["file1.xlsx", "file2.xlsx", "file3.xlsx"]  # Add your file names here
combined_df = pd.concat([pd.read_excel(file) for file in excel_files], ignore_index=True)

# Remove duplicates
combined_df.drop_duplicates(inplace=True)

# Outlook email configuration
outlook_email = "mahamsyed6171@example.com"
outlook_password = "typeyourpassword"

# SMTP server configuration
smtp_server = "smtp-mail.outlook.com"
smtp_port = 587

# Email template
subject = "Your Subject"
body = "Your email body here with {} placeholders for dynamic content."

# Function to send emails
def send_email(to_email, subject, body):
    msg = MIMEMultipart()
    msg['From'] = outlook_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(outlook_email, outlook_password)
        server.sendmail(outlook_email, to_email, msg.as_string())

# Send emails (limit to 300 per day)
daily_email_limit = 300
emails_sent_today = 0

for index, row in combined_df.iterrows():
    if emails_sent_today >= daily_email_limit:
        break

    to_email = row['email_column']  # Replace 'email_column' with the actual column name containing email addresses
    send_email(to_email, subject, body.format(row['dynamic_content_column']))  # Replace 'dynamic_content_column'

    emails_sent_today += 1

# Save progress or any necessary information for future runs