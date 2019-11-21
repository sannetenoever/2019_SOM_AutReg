# -*- coding: utf-8 -*-
"""
Created on Fri Nov  8 15:14:12 2019

Subfunction that give readwrite access for google sheets 
and read specific information out of sheet

USED:
https://developers.google.com/sheets/api/samples/reading
https://developers.google.com/sheets/api/quickstart/python

Install needed:
    pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

@author: sjgte
"""

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
from datetime import datetime

#%% start session for email
def openreadwriteEmail(credloc):
    SCOPESemail = ['https://www.googleapis.com/auth/gmail.send']

    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(credloc+'tokenEmail.pickle'):
        with open(credloc+'tokenEmail.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credloc+
                'credentialsEmail.json', SCOPESemail)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(credloc+'tokenEmail.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    return service

#%% start the session
def openreadwrite(credloc):    
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    if os.path.exists(credloc+'token.pickle'):
        with open(credloc+'token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credloc+'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(credloc+'token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    service = build('sheets', 'v4', credentials=creds)
    return service

#%% read values from specified range
def readandprint(sheet, SAMPLE_SPREADSHEET_ID,SAMPLE_RANGE_NAME, printv):
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
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

#%% get name of a specific sheet value
def getnamesh(sheetinfo,sheetnum = 0):
    nameSh1 = sheetinfo['sheets'][sheetnum]['properties']['title']
    return nameSh1    

#%% read excel data to a pandas data frame
def readtoDF(sheet, ID,nameSht):
    allc = readandprint(sheet,ID,nameSht+'!'+'A1:Z1000',0)
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

#%% check the column name (maybe irrelevant)
def checkcol(sheet, ID, nameSht, columntitle, rowv = 1):
    nms = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    val = readandprint(sheet,ID,nameSht+'!A' + str(rowv) + ':Z' + str(rowv),0)
    matches = [num for num, x in enumerate(val[0]) if x == columntitle]
    colinx = nms[matches[0]]
    return colinx
        
#%%
def getnow():
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    return dt_string    
    
    
    
    
    
    