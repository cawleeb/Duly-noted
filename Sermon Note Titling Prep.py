# -*- coding: utf-8 -*-
"""
Created on Thu Jun  1 01:02:02 2023

@author: caleb
"""

#%% OneNote Sermon Titling Automation Script
#  Basic appearance should be

# Date 
# Book chaper:verses
# Pastor Name
# Sermon Title

"""Note, yes this will be imperfect, and the sermons will not always
perfectly reflect the outline due to last minute changes, for example.
The point is to get some basic experience with automation and email though. 
Needs to loop each week, should probably just display the day of at midnight. 
Note that the email is inconsistent about coming to my regular inbox versus 
'promotions'

"""
import email
from email import message_from_string
from email import policy
from email.parser import BytesParser
from email.parser import HeaderParser
import pandas as pd
import imaplib
from datetime import datetime
import pyautogui
import time
# import schedule
    # Will just use Windows Task scheduler for now
import re

# 1 Get the creds, stored offsite. 
file_path = 'C:/Users/caleb/Documents/ShareX/Hurl.txt'
try:
    with open(file_path, 'r') as file:
        # Step 2: Read the content
        file_content = file.read()

        # Step 3: Parse the content (Example: Splitting lines)
        creds = file_content.split('\n')
        incoming_email = creds[1]
        PASSWORD = creds[2]
        my_email = creds[3]
except FileNotFoundError:
    print(f"The file '{file_path}' was not found.")
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    # Step 4: Close the file
    file.close()

# IMAP server credentials and connection settings
IMAP_SERVER = 'imap.gmail.com'
port = 993
# would GGC's email be used here?
# Need to research the security of a script like this that can be accessed by 
# someone from the outside. 
# Thankfully, not much of my matieral will be incentivized to be accessed 
# until I begin to work on more high profile projects.
# Accessing from the server seems like the best option though, since the 
# PC doesn't have to be on to have the email pulled
# Start with the basics and easiest possible approach, even if
# it's overkill, then trim down as I'd like. 
# Comments can be put wherever I like, just should be trimmed down if code 
# is ever public.
# I mostly overthink my approach, in that it will rarely be perfect on the 
# first testing attempt, but I need to start testing regardless.
# When in doubt, K.I.S.S, one step at a time
# Use it or lose it. I'm ending up not actually 
                        # incoming_email, my_email, PASSWORD, IMAP_SERVER
                        # Test some inputs   
                        
                        
                       
def get_email_headers(IMAP_SERVER, port, my_email, incoming_email, PASSWORD):
    # Just need to do the basics right now.
    # Start with getting the info from the email
    # Connect to the IMAP server
    imap = imaplib.IMAP4_SSL(IMAP_SERVER, port)
    imap.login(my_email, PASSWORD)
    imap.list()
    imap.select('inbox')
    # This should be progressed to be able to search both inboxes 
    result, data = imap.uid('search', None, "SEEN")
    i = len(data[0].split())  
    # print(i)
    
    # How should this loop be adjusted? 
    # Ought to be based on the most recent date...
    # 
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
            date = str(email.header.make_header(email.header.decode_header(email_message['Date'])))
            # print("The date from the string shows as: ", date)
            date_object = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
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
    

    # mailbox = 'INBOX' #, 'PROMOTIONS']
    #                 # Can add in promotions later
    # email_headers = []
    
    # # for folder in mailbox:
    # imap.select(mailbox)
    # status, email = imap.search(None, f'(FROM "{incoming_email}")') #, 'UID')
    #                             # Not sure if this should be kept of not ^
    # # email_id = email[0].split()
    # email_ids = email[0].split()
    # print("email id's are: ", email_ids)
    # # Is there anything else this search needs to include?
    # # Search for emails from the specified email address
    # # Should the loops be nested?
    # # Or should I use an `if` statement
    
    # # Iterate through the email IDs and fetch the headers
    # # for email_id in email_ids:
    # # _, header_data = imap.fetch(email_id,        
    # _, header_data = imap.fetch(email_ids, '(BODY[HEADER.FIELDS (HEADER)])', '(RFC822)')[1]
    # # How to align my code with the pre-requisites for the _.fetch() tool?
    # # Remember, I'm only using this function so far because it's what GPT 
    # # told me to use
    # print('email header data is: ', header_data, '\n')
    # # header_bytes = bytes('\r\n'.join(header_data), 'utf-8')
    # headers = email.message_from_bytes(header_bytes[0][1])

    #     imap.logout() # Close the connection
    # else:
    #     print("Wrong GGC email.")

    # return email_headers
    imap.logout()

get_email_headers(IMAP_SERVER, port, my_email, incoming_email, PASSWORD)


#%% Comment overview
    # if time >= Saturday after 11:30pm
        # check inbox and promotions for GGC email with the 
        # subject 'worship preview'
    # 
    # Date + 1 as the sermon date
    # Subject: Book as the passage
    # Pastor's name
    # Assign the subject as the sermon title

# Open a new onenote subpage
# Format and paste the above text in
# all done!

        
    


# Combine the useful elements of this function with the earlier
def get_email(email_message):
    # How to extract the specfic email here?
    """Extract plain text from an email message."""
    # if email_message.is_multipart():
    # Timing items can go outside the function    
    # if email is a worship preview from GGC
    # if 
    # was sent within the given timeframe
    for part in email_message.walk():
        # Indentify the content type here
        content_type = part.get_content_type()
        if content_type == 'from':
            # Determine if GGC sent the email or not
            return part.get_payload(decode=True).decode()
        if content_type == 'text/plain':
            return part.get_payload(decode=True).decode()
    # else: 
        """assuming this means that there's just text in the email"""
    #     return email_message.get_payload(decode=True).decode()

#%% might be out of order
# Get the current date and time
current_time = datetime.datetime.now()

# Assuming 'email_sent_time' is the timestamp of the email... Is this a good 
# assumption?
email_sent_time = datetime.datetime(2023, 1, 25, 10, 30)  
# Replace with the actual email sent time


# Assuming 'cutoff_time' is the desired cutoff time
cutoff_time = datetime.time(11, 0)  # Replace with the actual desired cutoff time

# Compare the time part of 'email_sent_time' with the 'cutoff_time'
# if email_sent_time.time() < cutoff_time:
#     print("Email was sent before the cutoff time.")
# else:
#     print("Email was sent after the cutoff timeP.")
#%%
import requests
from msal import ConfidentialClientApplication

# Set your access token and section ID
# access_token = "YOUR_ACCESS_TOKEN"
client_secret = "d5cab2b7-d47f-4bd0-a793-1ef2c163e02c"
client_id = "331f975d-97d6-452d-9a89-20e9221c8e4d"
section_id = "600e812f-df7a-4ee4-857c-7c0ab482e9a2"
tenant_id = "a8176dbe-1d0e-4083-92ae-a108d22052af"

# Create a ConfidentialClientApplication
app = ConfidentialClientApplication(
    client_id, authority=f"https://login.microsoftonline.com/{tenant_id}",
    client_credential=client_secret
)
# Acquire an access token
scopes = ["https://graph.microsoft.com/.default"] # Replace with your desired scope
result = app.acquire_token_silent(scopes=scopes, account=None)

if "access_token" in result:
    access_token = result["access_token"]
    print(f"Access token: {access_token}")
else:
    print("Error acquiring access token")

# Construct the request URI
url = f"https://graph.microsoft.com/v1.0/me/onenote/sections/{section_id}/pages"

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
