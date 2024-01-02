############################################################################
# Python Script created by Frits van Iperen and AI
# Date 01-01-2024
# Set in crontab every day at 9AM
############################################################################
# Short discription
############################################################################
# This script wil set the email SMTP server settings.
# Import data from excel file
# Verify the cuurent date with de date of birth in the data excel file.
# When there is a match between current date and Date of birth.
# A birthday message is automatically send to the receiver email adres
############################################################################
import pandas as pd
import datetime
import smtplib
import logging
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# Read the Excel sheet
data = pd.read_excel('birthday_data.xlsx')

# Set up email details
from_email = 'smtp server login username'
password = 'passwordstring'
smtp_server = 'smtp.server.name'
smtp_port = 587

# Log in to the SMTP server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(from_email, password)

# Iterate through the rows
for idx, row in data.iterrows():
    firstname = row['firstname']
    lastname = row['lastname']
    email = row['email']
    dob = row['dob']

    birthday_message = f"""
        Beste {firstname} {lastname},
       Type hier je bericht

        """
    # Create the email message
    message = MIMEMultipart()
    message['From'] = from_email
    message['To'] = email
    message['Subject'] = "Happy Birthday!"
# Check if it's the person's birthday today
if dob.month == datetime.date.today().month and dob.day == datetime.date.today().day:
    # Add the birthday message to the email body
    message.attach(MIMEText(birthday_message, 'plain'))
    # Send the email
    server.send_message(message)
#Set up the logging configuration
with open('dob.log', 'w') as f:
    # Redirect stdout to the log file
    sys.stdout = f
    print(f"Email sent to {email}.")
    sys.stdout = sys.__stdout__

# Disconnect from the SMTP server
server.quit()
