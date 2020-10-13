#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#General library modules
import pandas as pd


# In[ ]:


#Email procssing libraries
import imaplib
from imaplib import IMAP4
import email


# In[ ]:


#Text Processing Libraries
import re
from bs4 import BeautifulSoup as bs4
import chardet


# EMail Credentials

# In[ ]:


print('Gmail and Yahoo Mail services are provided')
mservice = input('Select Mail Service : ')
imap_user = input('Username : ')
imap_pass = input('Password : ')

if mservice == 'yahoo':
    imap_host = 'imap.mail.yahoo.com'
if mservice == 'gmail':
    imap_host =  'imap.gmail.com'


# In[ ]:


imap = imaplib.IMAP4_SSL(imap_host)
imap.login(imap_user, imap_pass)


# In[ ]:


print(imap.list())


# In[ ]:


mailbox = input('Enter the Mailbox to download : ')
imap.select(mailbox)


# In[ ]:


typ, data = imap.search(None, 'ALL')
mails = data[0].split() #splitting into byte list


# In[ ]:


def cleanHtml(soup):
    for s in soup(['script', 'style']):
            s.decompose()
    cleanSoup = ' '.join(soup.stripped_strings)
    return cleanSoup


# In[ ]:


def extract_msg(emails):

    #lists to store Date and email messages
    date_list = []
    msgText = []

    for item in emails:
        result2, email_data = imap.fetch(item, '(RFC822)')
        raw_email = email_data[0][1].decode('latin1')
        email_msg = email.message_from_string(raw_email)
    
        if email_msg.is_multipart(): 
            html = None
            for part in email_msg.get_payload():
                if part.get_content_type() == 'text/plain':
                    text = str(part.get_payload(decode=True),str('latin1'),"ignore").encode('utf8','replace')
            
                if part.get_content_type() == 'text/html':
                    soup = bs4(str(part.get_payload(decode=True),str('latin1'),"ignore").encode('utf8','replace'), 'lxml')
                    html = cleanHtml(soup)
            
                if part.get_content_type() == 'multipart/alternative':
                    for subpart in part.get_payload():
                        if subpart.get_content_charset() is None:
                            charset = chardet.detect(bytes(subpart))['encoding']
                        else:
                            charset = subpart.get_content_charset()

                    if subpart.get_content_type() == 'text/plain':
                        text = str(subpart.get_payload(decode=True),str('latin1'),"ignore").encode('utf8','replace')

                    if subpart.get_content_type() == 'text/html':
                        soup = bs4(str(subpart.get_payload(decode=True),str('latin1'),"ignore").encode('utf8','replace'), 'lxml')
                        html = cleanHtml(soup)

            if html is None:
                text.strip()
                msgText.append(text)
            else:
                html.strip()
                msgText.append(html)
        else:
            if email_msg.get_content_charset() is None:
                charset = chardet.detect(bytes(email_msg))['encoding']
            else:
                charset = email_msg.get_content_charset()
            text = bs4(str(email_msg.get_payload(decode=True), str('latin1'),'ignore').encode('utf8','replace'), 'lxml')
            text = cleanHtml(text)
            text.strip()
            msgText.append(text)
    imap.close()
    imap.logout()
    return msgText


# In[ ]:


mail = extract_msg(mails)


# In[ ]:


# creating pandas dataframe
filename = mailbox+'.csv'
df = pd.DataFrame(mail)
dfT = df.T


# In[ ]:


dfT.to_csv(filename, mode='a', index=False)

