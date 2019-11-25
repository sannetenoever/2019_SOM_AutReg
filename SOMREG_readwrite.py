# -*- coding: utf-8 -*-
"""
Created on Fri Nov  8 15:14:12 2019

Subfunction that give readwrite access for google sheets, google email, forms
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

#%% give access to email, spreadsheets and forms
def CredentialsMultiple(credloc):
    SCOPESa = ['https://www.googleapis.com/auth/gmail.send', 
               'https://www.googleapis.com/auth/spreadsheets']
                #'https://www.googleapis.com/auth/forms']

    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(credloc+'tokenA.pickle'):
        with open(credloc+'tokenA.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credloc+
                'credentials.json', SCOPESa)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(credloc+'tokenA.pickle', 'wb') as token:
            pickle.dump(creds, token)

    EmServ = build('gmail', 'v1', credentials=creds)
    ShServ = build('sheets','v4',credentials=creds)
    #FoServ = build('forms','v2', credentials=creds) at this moment forms doesn't work
    return EmServ, ShServ #FoServ

#%% start session for email, first time needs to get the credentials
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

#%% start the session for sheets
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


    
    
    
    
    