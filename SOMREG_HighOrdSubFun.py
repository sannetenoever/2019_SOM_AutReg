# -*- coding: utf-8 -*-
"""
Created on Fri Nov  8 16:18:52 2019

Do some higher order things with the sheet:
    - make a summary sheet
    - email 

https://developers.google.com/sheets/api/samples/reading

@author: sjgte
"""

#import SOMREG_readwrite as rw
import pandas as pd
import numpy as np
import SOMREG_LowOrdSubFun as lsf
from tkinter import messagebox

#%% dump columns back in the forms (what was added was at the end)
def dumpcolumns(alldata, SheetA, nameSht, rangest='A1'):
    # write in json format > easy for columns
    bod = {'values':[alldata]}
    
    SheetA['serv'].spreadsheets().values().update(spreadsheetId=SheetA['ID'], range= nameSht +  '!' + rangest + ':Z2000',
                          valueInputOption = 'RAW', body = bod).execute()

#%% dump what is currently in alldata back into the form responses
def dumpdata(alldata, SheetA, nameSht, rangest='A1'):
    # first do the columsn
    col = alldata.columns
    col = col.values.tolist()
    # writ in json formate
    listv = alldata.values.tolist()
    listv = [col] + listv
    bod = {'values':listv}  
    SheetA['serv'].spreadsheets().values().update(spreadsheetId=SheetA['ID'], range= nameSht +  '!' + rangest + ':Z2000',
                          valueInputOption = 'RAW', body = bod).execute()

#%% check if the form has the relevant columns and initialize orderwise
def checkFormAndInit(SheetA, nameSht, rowv=1):    
    val = lsf.readandprint(SheetA,nameSht+'!A' + str(rowv) + ':Z' + str(rowv),0)

    # check if columns are present
    NeededColumns = ['Registration Type','Email Address', 'Role(s)','Full Name']
    MissingColumns = []
    for colnm in NeededColumns:
        if (colnm in val[0]) == False:
            MissingColumns.append(colnm)
    
    if len(MissingColumns)>0:
        messagebox.showinfo('Title', 'Form Mistake, missing entries:' + '/'.join(MissingColumns) + '\n\n ALSO CHECK SPACES')
    else:       
        # add columns that are not in the form but manually added
        ColumnsToAddIfMissing = ['Confirmation sent?','Confirmed as','Waitlisted?','Bank or cash?','Amount','Notes']
        for colnm in ColumnsToAddIfMissing:
            if (colnm in val[0]) == False:
                val[0].append(colnm)
        # dump all back in form
        dumpcolumns(val[0], SheetA, nameSht)
        
#%% provides a summary sheet
def sumSheet(SheetA, alldata):
    # print basics:
    # new sheet if needed
    if lsf.getnamesh(SheetA['sheetinfo'],1) != 'Summary':
        req = {"requests": [{"addSheet": { "properties": {
          "title": "Summary", "gridProperties": {"rowCount": 20,"columnCount": 12}}}}]}        
        request = serv.spreadsheets().batchUpdate(spreadsheetId=SheetA['ID'], body=req)
        request.execute()  
    
    # summary for the registration
    Type = ['Lead', 'Follow','Couple','Ambi']
    Conf = ['yes', 'no']
    Val = ['Student', 'Regular','Other','All']    
    
    # write the basic body
    bod = {'values':[['Summary ' + lsf.getnow()],
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
    SheetA['serv'].spreadsheets().values().update(spreadsheetId=SheetA['ID'], range='Summary!A1:F20',
                          valueInputOption = 'RAW', body = bod).execute()
    
    # assumed that Ccnt is the highest level. The rests gets counted in... 
    alldata = lsf.readtoDF(SheetA, lsf.getnamesh(SheetA['sheetinfo'],0))
    
    val = {'values': [[] for i in range(20)]}
    for Ccnt,Cval in enumerate(Conf):
        C = alldata[alldata['Confirmation sent?'].str.contains(Cval)==True]
        for Tcnt, Tval in enumerate(Type):
            if Cval == 'yes':
                CT = C[C['Confirmed as'].str.contains(Tval)==True]
            else:
                CT = C[C['Role(s)'].str.contains(Tval)==True]
            for Vcnt, Vval in enumerate(Val):
                if Vval == 'All':
                    CTV = CT
                elif Vval == 'Other':
                    searchfor = ['Student','Regular']
                    CTV = CT[~CT['Registration Type'].str.contains('|'.join(searchfor))]
                else:
                    CTV = CT[CT['Registration Type'].str.contains(Vval)==True]                   
                val['values'][len(Val)*Ccnt+Vcnt] = val['values'][len(Val)*Ccnt+Vcnt] + [len(CTV)]                
    val['values'].insert(len(Val),[''])
    colv = lsf.checkcol(SheetA,'Summary',Type[0],2)     
    
    # for confirmed make summary of final amount of couples and extras:
    allconf = val['values'][Conf.index('yes')*4+Val.index('All')]
    leadfol = [allconf[Type.index('Lead')]]+[allconf[Type.index('Follow')]]
    lfname = ['Lead','Follow']
    index_min = np.argmin(leadfol)
    index_max = np.argmax(leadfol)    
    TotCop = leadfol[index_min]+allconf[Type.index('Couple')]
    TotAmbi = allconf[Type.index('Ambi')]
        
    val['values'][len(Val)*len(Conf)+len(Conf)].append('Totals Confirmed:')
    val['values'][len(Val)*len(Conf)+len(Conf)+1].append('Couples:')
    val['values'][len(Val)*len(Conf)+len(Conf)+1].append(TotCop)
    val['values'][len(Val)*len(Conf)+len(Conf)+2].append('Extra '+lfname[index_max])
    val['values'][len(Val)*len(Conf)+len(Conf)+2].append(leadfol[index_max]-leadfol[index_min])
    val['values'][len(Val)*len(Conf)+len(Conf)+3].append('Ambis')
    val['values'][len(Val)*len(Conf)+len(Conf)+3].append(TotAmbi)  
      
    relout = {'ConfirmedCouples': TotCop, 
              'Extras': leadfol[index_max]-leadfol[index_min], 'ExtraRole': lfname[index_max],
              'Ambis':TotAmbi}
    
    # write to excel sheet       
    SheetA['serv'].spreadsheets().values().update(spreadsheetId=SheetA['ID'], range='Summary!' + colv + '3:Z100',
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
    SheetA['serv'].spreadsheets().values().update(spreadsheetId=SheetA['ID'], range='Summary!I2:O20',
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
            inxT.append(list(alldata['Registration Type'].str.contains(Rval)==True)) 

        # get the amount that should have been paid (ONLY WORKS FOR INTEGERS NOW!)
        pv = list(alldata['Registration Type'][inxT[-1]])
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
    SheetA['serv'].spreadsheets().values().update(spreadsheetId=SheetA['ID'], range='Summary!J4:O20',
                          valueInputOption = 'RAW', body = val).execute()
        
    return relout  

#%% check and send the confirmation email and update sheet, send waiting list if needed
def checkRolesConfirmation(SheetA, nameSht, alldata, sumDat, EmServ):       
     dfsel = pd.DataFrame(alldata[alldata['Confirmation sent?'].str.contains('no') == True])
     # prefill confirmed as none
     dfsel.loc[:,'Confirmed as'] = 'none'
         
     # go row by row to check if there is space
     # ambis get assign at a later point, not yet     
     NoElLeft = False
     cnt = 0
     newSumDat = {}
     while NoElLeft == False and len(dfsel)>0:
         dfinx = dfsel.index
         # define couples including ambi couples
         newSumDat = {}
         newSumDat['ExtraRole'] = sumDat['ExtraRole']
         if (sumDat['Ambis'] < sumDat['Extras']):
             # not enough ambis to fill the extras
             newSumDat['Extras'] = sumDat['Extras']-sumDat['Ambis']
             newSumDat['ConfirmedCouples'] = sumDat['ConfirmedCouples']+sumDat['Ambis']
         elif (sumDat['Ambis'] == sumDat['Extras']):
             # same amount of amis as extras:
             newSumDat['Extras'] = 0
             newSumDat['ConfirmedCouples'] = sumDat['ConfirmedCouples']+sumDat['Ambis']
         elif (sumDat['Ambis'] > sumDat['Extras']):
             newSumDat['ExtraRole'] = 'Ambi'
             # more ambis than extras, the ambis themselves can form couples:
             ExtraAm = sumDat['Ambis']-sumDat['Extras']
             AmbiCoupl = int(np.floor(np.array(ExtraAm)/2))
             if ExtraAm % 2 == 1:
                 newSumDat['Extras']=1 # only extra if odd number of ambis
             else:
                 newSumDat['Extras']=0
             newSumDat['ConfirmedCouples'] = sumDat['ConfirmedCouples']+AmbiCoupl+sumDat['Extras']
             # possible that we reach to too many couples. Control for that
             if newSumDat['ConfirmedCouples']>sumDat['MaxCouples']:
                 newSumDat['Extras'] = newSumDat['Extras']+(newSumDat['ConfirmedCouples']-sumDat['MaxCouples'])*2
                 newSumDat['ConfirmedCouples'] = sumDat['MaxCouples']
         
         # go through single person
         inxV = dfinx[cnt]
         Role = dfsel['Role(s)'][inxV]
                    
         # class full, so no need to check anymore
         if (newSumDat['ConfirmedCouples'] >= sumDat['MaxCouples']) and (newSumDat['Extras'] >= sumDat['MaxExtras']):
             NoElLeft = True
             
         # go through all options
         if 'no' in dfsel['Confirmation sent?'][inxV]:             
            # couples
            if ('Couple' in Role) and (newSumDat['ConfirmedCouples'] < sumDat['MaxCouples']):
                 # couple can be added as there are less than max couples
                 sumDat['ConfirmedCouples'] = sumDat['ConfirmedCouples']+1
                 dfsel = lsf.changeConfirm(dfsel,inxV, 'Couple')
                 cnt = cnt+1 # go to next, as this doesn't help people earlier on the list
            elif ('Couple' in Role) and (newSumDat['ConfirmedCouples'] >= sumDat['MaxCouples']):
                 # couple cannot be added as max couples reached
                 dfsel = lsf.changeConfirm(dfsel,inxV, 'WL COUPLES FULL', True)
                 cnt = cnt+1 # go to next, as this doesn't help people earlier on the list
             
            # ambis
            elif ('Ambi' in Role) and (newSumDat['ConfirmedCouples'] < sumDat['MaxCouples']):
                # ambis can always join if the couples are not full
                sumDat['Ambis'] = sumDat['Ambis']+1
                dfsel = lsf.changeConfirm(dfsel,inxV, 'Ambi')
                cnt = 0 # this could help other people in list
            elif ('Ambi' in Role) and (newSumDat['ConfirmedCouples'] >= sumDat['MaxCouples']):
                # if couples full ambis can only join if the amount of extras is not yet full
                if newSumDat['Extras'] < sumDat['MaxExtras']:
                    sumDat['Ambis'] = sumDat['Ambis']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Ambi')
                    cnt = 0 # might have influence on amount of couples/extras
                else:
                     # if ambi cannot be added anymore, then nobody can be added anymore
                     dfsel = lsf.changeConfirm(dfsel,inxV, 'WL ALL FULL', True)
                     NoElLeft = True
            
            # follows
            elif ('Follow' in Role) and (newSumDat['ConfirmedCouples'] < sumDat['MaxCouples']):
                if (sumDat['ExtraRole'] == 'Lead') and (sumDat['Extras'] > 0):
                    # follow can be added as enough space for new couple + extra leads
                    sumDat['Extras'] = sumDat['Extras']-1
                    sumDat['ConfirmedCouples'] = sumDat['ConfirmedCouples']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Follow')
                    cnt = 0 # could be that in list was a leader that can now be added
                elif sumDat['Extras'] == 0:                    
                    # follow can be added as enough space for new couple + no extras
                    sumDat['Extras'] = 1
                    sumDat['ExtraRole'] = 'Follow'
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Follow')
                    cnt = 0 # could help an ambi
                elif (sumDat['ExtraRole'] == 'Follow') and (newSumDat['Extras']  < sumDat['MaxExtras']):                  
                    # follow can be added as not yet max amount of extra follows
                    sumDat['Extras'] = sumDat['Extras']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Follow')
                    cnt = 0 # could help an ambi
                else:
                    # follow cannot be added, max of extra follows reached
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'WL MAX EXTRA FOLLOW', True)
                    cnt = cnt+1
            elif ('Follow' in Role) and (sumDat['ConfirmedCouples'] >= sumDat['MaxCouples']):
                # couples are full, Follow can be added as extra
                if (newSumDat['Extras'] == 0):
                    sumDat['Extras'] = sumDat['Extras'+1]
                    sumDat['ExtraRole'] = 'Follow'
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Follow')
                    cnt = cnt + 1 # this doesn't help people earlier on the list
                elif (sumDat['ExtraRole'] == 'Follow') and (newSumDat['Extras'] < sumDat['MaxExtras']):
                    sumDat['Extras'] = sumDat['Extras']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Follow')                 
                    cnt = cnt+1 # this doesn't help people earlier on the list
                else:
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'WL MAX EXTRA FOLLOW', True)
                    # nobody can be added anymore:
                    NoElLeft = True
                    
            # leads
            elif ('Lead' in Role) and (newSumDat['ConfirmedCouples'] < sumDat['MaxCouples']):
                if (sumDat['ExtraRole'] == 'Follow') and (sumDat['Extras'] > 0):
                    # lead can be added as enough space for new couple + extra follows
                    sumDat['Extras'] = sumDat['Extras']-1
                    sumDat['ConfirmedCouples'] = sumDat['ConfirmedCouples']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Lead')
                    cnt = 0 # could be that in list was a follow that can now be added
                elif sumDat['Extras'] == 0:                    
                    # lead can be added as enough space for new couple + no extras
                    sumDat['Extras'] = 1
                    sumDat['ExtraRole'] = 'Lead'
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Lead')
                    cnt = 0 # could help an ambi
                elif (sumDat['ExtraRole'] == 'Lead') and (newSumDat['Extras']  < sumDat['MaxExtras']):                  
                    # lead can be added as not yet max amount of extra leads
                    sumDat['Extras'] = sumDat['Extras']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Lead')
                    cnt = 0 # could help an ambi
                else:
                    # lead cannot be added, max of extra lead reached
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'WL MAX EXTRA LEADS', True)
                    cnt = cnt+1
            elif ('Lead' in Role) and (sumDat['ConfirmedCouples'] >= sumDat['MaxCouples']):
                # couples are full, Lead can be added as extra
                if (newSumDat['Extras'] == 0):
                    sumDat['Extras'] = sumDat['Extras'+1]
                    sumDat['ExtraRole'] = 'Lead'
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Lead')
                    cnt = cnt + 1 # this doesn't help people earlier on the list
                elif (sumDat['ExtraRole'] == 'Lead') and (newSumDat['Extras'] < sumDat['MaxExtras']):
                    sumDat['Extras'] = sumDat['Extras']+1
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'Lead')
                    cnt = cnt+1 # this doesn't help people earlier on the list
                else:
                    dfsel = lsf.changeConfirm(dfsel,inxV, 'WL MAX EXTRA LEAD', True)
                    # nobody can be added anymore:
                    NoElLeft = True           
            else:
                cnt = cnt+1
         else:
            cnt = cnt+1    
         if cnt == len(dfinx):
             NoElLeft = True
            
     dfsel.loc[dfsel[dfsel['Confirmed as'].str.contains('none') == True].index,'Waitlisted?'] = 'WL FULL'
     return dfsel, sumDat, newSumDat
        
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
