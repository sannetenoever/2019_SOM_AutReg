# -*- coding: utf-8 -*-
"""
Created on Fri Nov 15 13:24:53 2019

@author: sjgte
"""
#path = 'C:/Users/sjgte/surfdrive/Python/API/GoogleSheet'
path = 'C:/Users/sjgte/surfdrive/Python\API\SOM_AutReg\\'
credloc = 'C:/Users/sjgte/surfdrive/Python\API\Credentials\\'
import os
os.chdir(path)
import SOMREG_HighOrdSubFun as st
import SOMREG_LowOrdSubFun as lst
import SOMREG_readwrite as rw
import SOMREG_SubGUI as gui

#%% initializing things
# ask user for input to start:
# solo or couples dance
#   if solo then max amount of people
#   if couples then max amount of couples + max amount of extras
# amount of teachers
# starting date of the class
# name of the class
# time of the class
# location of the class
# associated excel file
SOMREGapp = gui.app()


EmServ, ShServ = rw.CredentialsMultiple(credloc)

sheetID = '19aQ21jARdZL9ksMs-8i03na1MmLXppn7RRL733g2zx0'
MaxCouples = 20
Extras = 3

#% Spreadsheet initialization
sheet = ShServ.spreadsheets()
sheetinfo = sheet.get(spreadsheetId=sheetID).execute()
# put all info about servers and stuff in a dictionary
SheetA = {'serv': ShServ, 'sheet':sheet, 'sheetinfo': sheetinfo, 'ID':sheetID}

# Check columns of the form and change if wrong
st.checkFormAndInit(SheetA, lst.getnamesh(SheetA['sheetinfo'],0))

# read all data into dataframe
alldata = lst.readtoDF(SheetA, lst.getnamesh(sheetinfo,0))

# make a summary sheet
sumDat = st.sumSheet(SheetA, alldata)
sumDat['MaxCouples'] = MaxCouples
sumDat['MaxExtras'] = Extras

# check role inbalance and confirm person as X or Y
# still needs some final checking!!!
[dfsel, sumDat, newSumDat] = st.checkRolesConfirmation(SheetA, lst.getnamesh(sheetinfo,0), alldata, sumDat,EmServ)
alldata.loc[dfsel.index] = dfsel

# put updated data back in the excel sheet
st.dumpdata(alldata, SheetA, lst.getnamesh(sheetinfo,0))

# make sum sheet again
st.sumSheet(SheetA, alldata)

