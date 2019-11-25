# -*- coding: utf-8 -*-
"""
Created on Fri Nov 22 09:19:26 2019

Subfunction related to reading and writing things from sheets


@author: sjgte
"""
import pandas as pd
from datetime import datetime
from email.mime.text import MIMEText
import base64

#%% read values from specified range from sheet
def readandprint(SheetA,SAMPLE_RANGE_NAME, printv):
    result = SheetA['sheet'].values().get(spreadsheetId=SheetA['ID'],
                                range=SAMPLE_RANGE_NAME).execute()    
    values = result.get('values', [])

    if not values:
        print('No data found.')
    elif printv:
        #print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[1]))
    return values

#%% get title of a specific sheet
def getnamesh(sheetinfo,sheetnum = 0):
    nameSh1 = sheetinfo['sheets'][sheetnum]['properties']['title']
    return nameSh1    

#%% read excel data from full sheet to a pandas data frame
def readtoDF(SheetA,nameSht):
    allc = readandprint(SheetA,nameSht+'!'+'A1:Z1000',0)
    # put all the data in a panda frame
    for c in range(1,len(allc)):
        # check if the list is the correct length:
        lend = len(allc[0])-len(allc[c])
        toaddlist = []
        if lend > 0:
            toaddlist = allc[c]
            toaddlist.extend([0]*lend)
        else:
            toaddlist = allc[c][0:len(allc[0])]   
        # panda dataframe of a single entry
        pdfr = pd.DataFrame([toaddlist], index = [c-1], columns = allc[0])
        if c == 1:
            alldata = pdfr
        else:
            alldata = alldata.append(pdfr)
    return alldata

#%% check the column name (maybe irrelevant later)
def checkcol(SheetA, nameSht, columntitle, rowv = 1):
    nms = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    val = readandprint(SheetA,nameSht+'!A' + str(rowv) + ':Z' + str(rowv),0)
    matches = [num for num, x in enumerate(val[0]) if x == columntitle]
    colinx = nms[matches[0]]
    return colinx
        
#%% return the current date
def getnow():
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    return dt_string    
    
#%% send email
def create_message(to, subject, message_text):
  """Create a message for an email.

  Args:
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

#%% update the confirmation or put in waitlist
def changeConfirm(dfsel,inxV, role, WL=False):
    if WL == False:
        dfsel.loc[inxV,'Confirmation sent?'] = 'yes ' + getnow()
        dfsel.loc[inxV,'Confirmed as'] = role
    else:
        dfsel.loc[inxV,'Waitlisted?'] = role
        dfsel.loc[inxV,'Confirmation sent?'] = 'no'
    
    return dfsel


