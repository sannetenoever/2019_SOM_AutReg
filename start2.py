# -*- coding: utf-8 -*-
"""
Created on Fri Nov 15 13:24:53 2019

@author: sjgte
"""
#path = 'C:/Users/sjgte/surfdrive/Python/API/GoogleSheet'
path = 'D:\Surfdrive\Python\API\SOM_AutReg'
credloc = 'D:\Surfdrive\Python\API\Credentials\\'

import os
os.chdir(path)
import SubFunSheets as st
import readwrite as rw

#%% initializing things
ID = '19aQ21jARdZL9ksMs-8i03na1MmLXppn7RRL733g2zx0'

#% open the access to the file
serv = rw.openreadwrite(credloc)

#% check what the first
sheet = serv.spreadsheets()
sheetinfo = sheet.get(spreadsheetId=ID).execute()
nameSh1 = rw.getnamesh(sheetinfo,0) # name of the main first sheet
alldata = rw.readtoDF(sheet, ID, rw.getnamesh(sheetinfo,0))

#%%
st.sumSheet(serv,sheet,sheetinfo, alldata,ID)

