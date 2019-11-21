# -*- coding: utf-8 -*-
"""
Created on Fri Nov  8 16:18:52 2019

https://developers.google.com/sheets/api/samples/reading

@author: sjgte
"""

import readwrite as rw
import pandas as pd
import numpy as np

#%% provides a summary sheet
def sumSheet(serv,sheet,sheetinfo, alldata,ID):
    # print basics:
    # new sheet if needed
    if rw.getnamesh(sheetinfo,1) != 'Summary':
        req = {"requests": [{"addSheet": { "properties": {
          "title": "Summary", "gridProperties": {"rowCount": 20,"columnCount": 12}}}}]}        
        request = serv.spreadsheets().batchUpdate(spreadsheetId=ID, body=req)
        request.execute()  
    
    # summary for the registration
    Type = ['Lead', 'Follow','Couple','Ambi']
    Conf = ['yes', 'no']
    Val = ['Student', 'Regular','Other','All']    
    
    # write the basic body
    bod = {'values':[['Summary ' + rw.getnow()],
        ['','']+Type,
        ['Confirmed']+[Val[0]],
        ['']+[Val[1]],
        ['']+[Val[2]],
        ['']+[Val[3]],
        [''],
        ['NonConfirmed']+[Val[0]],
        ['']+[Val[1]],
        ['']+[Val[2]],
        ['']+[Val[3]]]}
    serv.spreadsheets().values().update(spreadsheetId=ID, range='Summary!A1:F20',
                          valueInputOption = 'RAW', body = bod).execute()
    
    # assumed that Ccnt is the highest level. The rests gets counted in... 
    alldata = rw.readtoDF(sheet, ID, rw.getnamesh(sheetinfo,0))
    
    val = {'values': [[] for i in range(20)]}
    for Ccnt,Cval in enumerate(Conf):
        C = alldata[alldata['Confirmation sent?'].str.contains(Cval)==True]
        for Tcnt, Tval in enumerate(Type):
            if Cval == 'yes':
                CT = C[C['Confirmed As'].str.contains(Tval)==True]
            else:
                CT = C[C['Role(s)'].str.contains(Tval)==True]
            for Vcnt, Vval in enumerate(Val):
                if Vval == 'All':
                    CTV = CT
                elif Vval == 'Other':
                    searchfor = ['Student','Regular']
                    CTV = CT[~CT['Registration Type '].str.contains('|'.join(searchfor))]
                else:
                    CTV = CT[CT['Registration Type '].str.contains(Vval)==True]                   
                val['values'][len(Val)*Ccnt+Vcnt] = val['values'][len(Val)*Ccnt+Vcnt] + [len(CTV)]                
    val['values'].insert(len(Val),[''])
    colv = rw.checkcol(sheet,ID,'Summary',Type[0],2)     
    
    # for confirmed make summary of final amount of couples and extras:
    allconf = val['values'][Conf.index('yes')*4+Val.index('All')]
    leadfol = [allconf[Type.index('Lead')]]+[allconf[Type.index('Follow')]]
    lfname = ['Lead','Follow']
    index_min = np.argmin(leadfol)
    index_max = np.argmax(leadfol)    
    TotCop = leadfol[index_min]*2+allconf[Type.index('Couple')]
        
    val['values'][len(Val)*len(Conf)+len(Conf)].append('Totals Confirmed:')
    val['values'][len(Val)*len(Conf)+len(Conf)+1].append('Couples:')
    val['values'][len(Val)*len(Conf)+len(Conf)+1].append(TotCop)
    val['values'][len(Val)*len(Conf)+len(Conf)+2].append('Extra '+lfname[index_max])
    val['values'][len(Val)*len(Conf)+len(Conf)+2].append(leadfol[index_max]-leadfol[index_min])
      
    # write to excel sheet       
    serv.spreadsheets().values().update(spreadsheetId=ID, range='Summary!' + colv + '3:Z100',
        valueInputOption = 'RAW', body=val).execute()
 
    ## summary for the financing    
    Reg = ['Student','Regular','Other','All']    
    Payment = ['B','C','NONE']

    # write the basic body
    bod = {'values':[['Numbers'],
        ['']+Payment,
        [Reg[0]],
        [Reg[1]],
        [Reg[2]],
        [Reg[3]],
        [''],
        [''],
        ['Amounts'],
        ['']+Payment,
        [Reg[0]],
        [Reg[1]],
        [Reg[2]],
        [Reg[3]]]}
    serv.spreadsheets().values().update(spreadsheetId=ID, range='Summary!I2:O20',
                          valueInputOption = 'RAW', body = bod).execute()
    
    inxT = []
    Payv = []
    TotAmount = []
    val = {'values': [[] for i in range(20)]}
    for Rcnt,Rval in enumerate(Reg):
        if Rval == 'Other': # nondefined
            inxT.append(list(np.array(inxT).sum(0)==0))
        elif Rval == 'All':
            inxT.append([True for it in range(0,len(inxT[-1]))])
        else:
            inxT.append(list(alldata['Registration Type '].str.contains(Rval)==True)) 

        # get the amount that should have been paid (ONLY WORKS FOR INTEGERS NOW!)
        pv = list(alldata['Registration Type '][inxT[-1]])
        if len(pv)>0:
            numinx = [char for char in pv[0] if char.isdigit()==True]
            if len(numinx)>0:
                Payv.append(int(''.join(numinx)))
            else:
                Payv.append('?')
        else:
            Payv.append('?')
        for Pcnt, Pval in enumerate(Payment):    
            if Pval == 'NONE':
                searchfor = ['B','C']
                inxB = alldata['Bank or cash?'].str.contains('|'.join(searchfor))==True
                inxB = inxB==False
            else:                
                inxB = alldata['Bank or cash?'].str.contains(Pval)==True
            fi = [i for i,j in enumerate(inxT[-1]) if j==True & inxB[i]==True]       
            # check amount of Couples in this list
            amC = sum(alldata['Role(s)'][fi].str.contains('Couple')==True)       
            TotPeople = len(fi)+amC
            
            ar = list(alldata['Amount'][fi])
            if len(ar)==1:
                try:
                    float(ar)
                    arF = [float(ar),0]
                except:
                    arF = [0,0]    
            elif len(ar)> 1:
                arIx = np.char.isnumeric(ar)
                arF = [float(i) for indx,i in enumerate(ar) if arIx[indx] == True]
            else:
                arF = [0,0]
                
            TotAmount.append(sum(arF))
            val['values'][Rcnt].append(float(TotPeople)) 
            val['values'][len(Reg)+Rcnt+4].append(TotAmount[-1])            
        val['values'][len(Reg)*2+5] = ['Total', sum(TotAmount)/2] # divided by to as also all is added            
    serv.spreadsheets().values().update(spreadsheetId=ID, range='Summary!J4:O20',
                          valueInputOption = 'RAW', body = val).execute()
          

#%% dump what is currently in alldata back into the form responses
def dumpdate(alldata, sheet):
    x=3

#%% check and send the confirmation email and update sheet
def confemail(sheet, ID, nameSht, alldata):
     dfsel = alldata[alldata['Confirmation sent?'].str.contains('no') == True]
     allinx = dfsel['Email Address']
     for cnt,sl in enumerate(dfsel):
         x = sl

#%% random code not used atm
# =============================================================================
# RANGE = 'Sheet1!B2'
# serv = openreadwrite()
# readandprint(serv, SPREADSHEET_ID, 'A1:B2')
# 
# #v = serv.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE).execute()
# bod = {"values": [[ 'getting wijntjes'], []]}
# vn = serv.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=RANGE, 
#                                valueInputOption = 'RAW', body = bod).execute()
#             
# readandprint(serv, SPREADSHEET_ID, 'A1:B2')
# =============================================================================
