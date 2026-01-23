# -*- coding: utf-8 -*-
"""
Created on Thu Jun  1 01:02:02 2023

@author: caleb
"""

#%% OneNote Sermon Titling Automation Script
#  Basic appearance ought to be

# Date 
# Book chaper:verses
# Pastor Name
# Sermon Title

import email
from email import message_from_string
from email import policy
from email.parser import BytesParser
from email.parser import HeaderParser
import pandas as pd
import imaplib
from datetime import datetime
# import pyautogui
import time
import requests
from msal import ConfidentialClientApplication
# from config import 
# import re
                     
                       
def get_email_headers(EMAIL_SERVER, EMAIL_PORT, RECEIVING_ADDRESS, INCOMING_EMAIL, EMAIL_PASSWORD):
    # Connect to the IMAP server
    imap = imaplib.IMAP4_SSL(EMAIL_SERVER)
    imap.port = EMAIL_PORT
    imap.login(RECEIVING_ADDRESS, EMAIL_PASSWORD)
    imap.list()
    imap.select('inbox')
    # label = 'new-lecture-note-page'
    # This should be progressed to be able to search both inboxes 
    result, data = imap.uid('search', None, "SEEN")
    i = len(data[0].split())  
    # print(i)
    
    # How should this loop be adjusted? 
    # Ought to be based on the most recent date...
   
    # Use the existing 'new-onenote-sermon-page' filter.
    # This ought to work best in thory, but if I 
    for x in range(i-30, i):
        
        latest_email_uid = data[0].split()[x]
        # print(latest_email_uid)
        result, email_data = imap.uid('fetch', latest_email_uid, '(RFC822)')
        # print(result)
        # print(email_data)
        
        raw_email = email_data[0][1]
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        # Header Details
        date_tuple = email.utils.parsedate_tz(email_message['Date'])
        if date_tuple:
            # local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
            local_date = datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
            local_message_date = "%s" %(str(local_date.strftime("%a, %d %b %Y"))) #" %H:%M:%S")))
        email_from = str(email.header.make_header(email.header.decode_header(email_message['From'])))
        # print("The subject line```` is: ")
        subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
        # print(email_from)
        if email_from == INCOMING_EMAIL and "Worship Preview:" in subject:
            # Get the current date and time
            current_time = datetime.datetime.now()
            # date = str(email.header.make_header(email.header.decode_header(email_message['Date'])))
            # print("The date from the string shows as: ", date)
            date_object = current_time.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
            # Can slice and dice up the string, but there's likely some 
            # functions I can find in Pandas or the like that will just
            # auto-convert it for me. 
            # this seems to be ignored in the terminal
            Date = date_object.strftime('%m/%d/%y')
            print(f"The formatted date shows as: {Date}")
            # Want to extract the pastor name as well
            # search email for text = phrase, "Bold"\
            # print()
            # print(f'The subject line pre-trim is {subject}')
            subject1 = subject.lstrip("Worship Preview: ")
            # Why is the stripping function getting rid of the final character?
            # print(f'The trimmed subject line is: {subject1}')
            Passage, Title = subject1.split(" | ")
            # print(f'The passage and title are: ')
            # print(Passage)
            # print(Title)
            Outline_strings = [Date,Passage,Title]
            Notepage_Title = "\n".join(Outline_strings)
            print(f'The note page title shows as: {Notepage_Title}')
            
            # pattern = re.compile(f'{re.escape(search_text)}(.+)', re.DOTALL)
    imap.logout()

get_email_headers(EMAIL_SERVER,EMAIL_PORT, RECEIVING_ADDRESS, INCOMING_EMAIL, EMAIL_PASSWORD)

#%% Combine the useful elements of this function with the earlier
def MakeNotePage(email_message):
    # How to extract the specfic email here?
    """Formatted email text for page outline inserted """
    # Set your access token and section ID
    # access_token = "YOUR_ACCESS_TOKEN"

    # Create a ConfidentialClientApplication
    app = ConfidentialClientApplication(
        CLIENT_ID, authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET
    )
    # Acquire an access token
    scopes = ["https://graph.microsoft.com/.default"] # Replace with your desired scope
    result = app.acquire_token_silent(scopes=scopes, account=None)

    if "access_token" in result:
        access_token = result["access_token"]
        # print(f"Access token: {access_token}")
    else:
        print(f"Error acquiring access token") #. Status code: {}")
        

    # Construct the request URI
    url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{SECTION_ID}/pages"

    # Define the input HTML (content of your subpage)
    input_html = "<div><p>04-26</p></div>"

    # Send the POST request
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/xhtml+xml",
    }
    response = requests.post(url, headers=headers, data=input_html)

    if response.status_code == 201:
        print("Subpage created successfully!")
    else:
        print(f"Error creating subpage. Status code: {response.status_code}")

# %%
