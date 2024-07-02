#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Mon Jul  1 22:01:07 2024

@author: fabian
"""

import smtplib
from email.message import EmailMessage
import pandas as pd
import logging
from dotenv import load_dotenv
import os
from datetime import datetime
import configparser
import time
import sys
from string import Template

# Load environment variables from .env file
load_dotenv()
password = os.getenv('EMAIL_PASSWORD')

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(message)s')

# Read SMTP settings and debug flag from settings.ini
config = configparser.ConfigParser()
config.read('settings.ini')

smtp_server = config['SMTP']['server']
smtp_port = int(config['SMTP']['port'])
login = config['SMTP']['login']
send_email_flag = config.getboolean('EMAIL', 'send_email_flag')
batch_size = int(config['EMAIL']['batch_size'])
pause_duration = int(config['EMAIL']['pause_duration'])
bcc_sender = config.getboolean('EMAIL', 'bcc_sender')

smv_logo_url = config['LOGOS']['smv_logo_url']
kaufleute_logo_url = config['LOGOS']['kaufleute_logo_url']
langendorf_logo_url = config['LOGOS']['langendorf_logo_url']

subject = config['BODY']['subject']
template_file = config['BODY']['template_file']

greetings_kaufleute = config['GREETINGS']['kaufleute']
greetings_langendorf = config['GREETINGS']['langendorf']

# Read and prepare the email body template using Template
with open(template_file, 'r', encoding='utf-8') as file:
    body_template = Template(file.read())

def send_email(to_email, subject, body, smtp_server, smtp_port, login, password, send_email_flag):
    msg = EmailMessage()
    msg.set_content(body, subtype='html')
    msg['Subject'] = subject
    msg['From'] = login
    msg['To'] = to_email
    
    if bcc_sender:
        msg['Bcc'] = login  # include BCC
    
    if send_email_flag:
        # Send the email
        try:
            if smtp_port == 465:
                with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                    server.login(login, password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()  # Upgrade the connection to a secure encrypted SSL/TTLS connection
                    server.login(login, password)
                    server.send_message(msg)
            logging.info(f"Email sent to {to_email}")
            return True
        except Exception as e:
            logging.error(f"Failed to send email to {to_email}: {e}")
            return False
    else:
        logging.info(f"Email prepared for {to_email} (debug mode, not sent)")
        # Save the email content to a file for debugging
        with open("example_email_sent.html", "w", encoding="utf-8") as f:
            f.write(body)
        return True

# Load the Excel file
data = pd.read_excel("openhelpers.xlsx")

# Specify the groups to send emails to
groups_send = ["Langendorf", "Kaufleute"]

# Counters for the number of emails sent
email_count = {
    "Langendorf": 0,
    "Kaufleute": 0
}

# Define log file name
log_filename = "sent_emails_log.xlsx"

# Load the existing log if it exists
if os.path.exists(log_filename):
    sent_emails_log_df = pd.read_excel(log_filename)
else:
    sent_emails_log_df = pd.DataFrame(columns=data.columns.tolist() + ['Sent Date'])

# Convert log DataFrame to a set for quick lookup
sent_emails_log_set = set(sent_emails_log_df['E-Mail']) if not sent_emails_log_df.empty else set()

# Loop through the data and send emails
batch_count = 0

for index, row in data.iterrows():
    gruppe = row['Gruppen']
    email = row['E-Mail']

    # Check if email has already been sent
    if email in sent_emails_log_set:
        logging.info(f"Email to {email} already sent, skipping...")
        continue

    if any(group in gruppe for group in groups_send):
        vorname = row['Vorname']
        nachname = row['Nachname']
        ist = row['Erreichter Wert in Stunden']
        soll = row['Zielwert in Stunden']
        
        # Determine the group logo and greetings based on the group
        if "Kaufleute" in gruppe:
            group_logo = kaufleute_logo_url
            greetings = greetings_kaufleute
            email_count["Kaufleute"] += 1
        elif "Langendorf" in gruppe:
            group_logo = langendorf_logo_url
            greetings = greetings_langendorf
            email_count["Langendorf"] += 1
        else:
            group_logo = ""  # Fallback logo if needed
            greetings = ""

        # Substitute placeholders in the body template
        body = body_template.safe_substitute(
            name=vorname, 
            ist=ist, 
            soll=soll, 
            group_logo=group_logo, 
            greetings=greetings, 
            smv_logo_url=smv_logo_url
        )
        
        # Log the details
        logging.info(f"Vorname: {vorname}")
        logging.info(f"Nachname: {nachname}")
        logging.info(f"Email: {email}")
        logging.info(f"IST/SOLL: {ist}/{soll}")
        logging.info(f"Gruppe: {gruppe}")
        
        email_sent = send_email(email, subject, body, smtp_server, smtp_port, login, password, send_email_flag)
        if email_sent:
            # Log the sent email immediately with sent date
            sent_emails_log_set.add(email)
            row['Sent Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            sent_emails_log_df = pd.concat([sent_emails_log_df, pd.DataFrame([row])], ignore_index=True)
            sent_emails_log_df.to_excel(log_filename, index=False)
        logging.info("-" * 40)
        
        batch_count += 1
        if batch_count >= batch_size:
            if index + 1 < len(data):
                logging.info(f"Pausing for {pause_duration} seconds to avoid SMTP server overload...")
                for remaining in range(pause_duration, 0, -1):
                    sys.stdout.write(f"\rResuming in {remaining} seconds...")
                    sys.stdout.flush()
                    time.sleep(1)
                logging.info("\nResuming now.")
                logging.info("-" * 40)
            batch_count = 0

# Summary logging
logging.info("=" * 40)
logging.info("Summary of emails sent:")
logging.info(f"Emails sent to Langendorf: {email_count['Langendorf']}")
logging.info(f"Emails sent to Kaufleute: {email_count['Kaufleute']}")
logging.info("=" * 40)

















