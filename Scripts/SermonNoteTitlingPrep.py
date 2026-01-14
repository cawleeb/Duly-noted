# -*- coding: utf-8 -*-
"""
Created on Thu Jun  1 01:02:02 2023

@author: caleb
"""

# OneNote Sermon Titling Automation Script
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
    imap = imaplib.IMAP4_SSL(EMAIL_SERVER, )
    imap.login(RECEIVING_ADDRESS, EMAIL_PASSWORD)
    imap.list()
    imap.select('inbox')
    # This should be progressed to be able to search both inboxes 
    result, data = imap.uid('search', None, "SEEN")
    i = len(data[0].split())  
    # print(i)
    
    # How should this loop be adjusted? 
    # Ought to be based on the most recent date...
   
    # Use the existing 'new-onenote-sermon-page' filter.
    for x in range(i-150, i):
        
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
        # print("The subject line is: ")
        subject = str(email.header.make_header(email.header.decode_header(email_message['Subject'])))
        # print(email_from)
        if email_from == 'Gospel Grace Church <info@gospelgrace.com>' and "Worship Preview:" in subject:
            # Get the current date and time
            current_time = datetime.datetime.now()
        #  date = str(email.header.make_header(email.header.decode_header(email_message['Date'])))
            # print("The date from the string shows as: ", date)
            date_object = current_time.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
            # Can slice and dice up the string, but there's likely some 
            # functions I can find in Pandas or the like that will just
            # auto-convert it for me. 
            # this seems to be ignored in the terminal
            email_headers = date_object.strftime('%m/%d/%y')
            print(f"The formatted date shows as: {email_headers}")
            # Want to extract the pastor name as well
            # search email for text = phrase, "Bold"\
            # print()
            # print(f'The subject line pre-trim is {subject}')
            subject1 = subject.lstrip("Worship Preview: ")
            # Why is the stripping function getting rid of the final character?
            # print(f'The trimmed subject line is: {subject1}')
            Passage, Title = subject1.split("|")
            print('The passage and title are: ')
            print(Passage)
            print(Title.lstrip(' '))
            
            # pattern = re.compile(f'{re.escape(search_text)}(.+)', re.DOTALL)

            # # Search for the pattern in the email body
            # match = pattern.search(email_body)
        
            # if match:
            #     # Extract the immediately continued line
            #     parsed_line = match.group(1).strip()
            #     return parsed_line
            # else:
            #     return None
    
            # _, msg_data = email.fetch(messages[0], '(RFC822)')
            # msg = mail.message_from_bytes(msg_data[0][1])
            # Should I use the .fetch() function here, or is that 
            # overcomplicating the needs of the program?
            
            # msg = email.message_from_bytes(msg_data[0][1])
            # for part in msg.walk():
            #     if part.get_content_type() == 'text/plain':
            #         body_text = part.get_payload(decode=True)
            # name = email.body_text.split('|')[1]
            
            
            # Perhaps find a way to designate the text component of the email 
            # as 'body_text' or the like
            # Pastor = name.splitlines()[0]
            # print(f"The pastor's name is {Pastor}")
            # print()
            # break
        
            
            
            # How to find the pastor name? 
            # Might not be any winning here, i.e. multiple pastors names could 
            # be in the email.
            # If they change up how the email is arranged, which seems to 
            # occasionally happen 
            # likely easier to implement this now, rather than latter and need 
            # to integrate into the existing function.
            # for snippet in email_body
                # if email_message.find('|') == true
                    #_, Elder = segment[snippet].split('|')
                    
                    # Keep in mind that many different people will preach at 
                    # GGC through the course of time, and thus this may not 
                    # always be a feasible part to work with the
            # Arrange to get the appearance: 
            # Date 08/06/23
            # Book chaper:verses
            # Pastor Name
            # Sermon Title
            
            
            # email_headers.append(Passage, Pastor, Title)
            
            
            # print(f"The title string appears as: {email_headers}")
                # 'Passage': headers.get('Subject'),
                # 'mmddyy': headers.get('mmddyy')
            # Subject = headers.get('Subject')
            # pattern = r":\s*(.*?)\s*\|\s*(.*?)$"
            # matches = re.search(pattern, Subject)

            # if matches:
            #     scripture = matches.group(1)
            #     sermon_title = matches.group(2)

            #     print(scripture)
            #     print(sermon_title)
            
            # for part in email_message.walk():
            #     if part.get_content_type() == "text/plain":
            #         body = part.get_payload(decode=True)
            #         print(body.decode('utf-8'))
            #         print()
            #         print()
            #         print()
            #         print()

                # else:
                #     continue
            # body = email.message.
    imap.logout()

get_email_headers(EMAIL_SERVER,EMAIL_PORT, RECEIVING_ADDRESS, incoming_email, EMAIL_PASSWORD)

# Combine the useful elements of this function with the earlier
def MakeNotePage(email_message):
    # How to extract the specfic email here?
    """Formatted email text for page outline inserted """

# Assuming 'email_sent_time' is the timestamp of the email... Is this a good 
# assumption?
# email_sent_time = datetime.datetime(2023, 1, 25, 10, 30)  
# Replace with the actual email sent time


# Assuming 'cutoff_time' is the desired cutoff time
# cutoff_time = datetime.time(11, 0)  # Replace with the actual desired cutoff time

# Compare the time part of 'email_sent_time' with the 'cutoff_time'
# if email_sent_time.time() < cutoff_time:
#     print("Email was sent before the cutoff time.")
# else:
#     print("Email was sent after the cutoff timeP.")

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
