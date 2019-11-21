# -*- coding: utf-8 -*-
"""
Created on Sun Nov 17 15:18:55 2019

@author: sjgte
"""
#%%
import readwrite as rw
servE = rw.openreadwriteEmail(credloc)

from email.mime.text import MIMEText
import base64

#% try to send email
def create_message(to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(message_text)
  message['to'] = to
  message['subject'] = subject
  
  #return message
  raw = base64.urlsafe_b64encode(message.as_bytes())
  raw = raw.decode()
  body = {'raw': raw}
  return body

#%%
to = 'mehrdad@seirafi.net'
subject = 'I will get some foodies'
message_text = 'Anything you want?'  

x = create_message(to, subject, message_text)
message = (servE.users().messages().send(userId='me', body=x)
               .execute())

