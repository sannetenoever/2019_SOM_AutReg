# -*- coding: utf-8 -*-
"""
Created on Fri Nov 22 11:40:21 2019

@author: sjgte
"""
#%%
import SOMREG_HighOrdSubFun as st

from tkinter import *
import tkinter as ttk
#from ttk import *

# ask user for input to start:

# 1) name of the class         
# 2) starting date             7) solo or couples dance
# 3) amount of weeks           8) max amount of people/couples
# 4) time of the class         9) if couples then max max amount of extras
# 5) location of the class     10) associated excel file ID
# 6) amount of teachers

#%%
root = Tk()
root.title("Initialization information")

# Add a grid
colmN = 2
rowN = 12

mainframe = Frame(root)
mainframe.grid(column=colmN,row=rowN, sticky=(N,W,E,S) )
mainframe.columnconfigure(0, weight = 1)
mainframe.rowconfigure(0, weight = 1)
mainframe.pack(pady = 100, padx = 100)

choices = ['Solo','Couples']
entrnames = {'Name of Course': ['entry','','string'],
             'Starting date': ['entry','','date'],
             'Amount of weeks': ['entry','','int'],
             'Time of class': ['entry','','time'],
             'Location of class': ['entry','string'],
             'Amount of teachers': ['entry','','int'],
             'Type of class': ['drop', choices[0], set(choices), 'string'],
             'Max N of couples/students':['entry', '', 'int'],
             'Max N of extras':['entry','','int'],
             'Google Sheet ID':['entry','','googlesheetID'],
             'OK':['button','OK','string']}

## 
listout = []
cntrow = 1
cntcol = 1
for it in entrnames:
    # text box
    listout.append(StringVar(root))
    if entrnames[it][0] == 'entry':
        Label(mainframe, text= it).grid(row = cntrow, column = cntcol)
        entry = ttk.Entry(mainframe, textvariable = listout[-1], font=16)
        entry.grid(row = cntrow+1, column = cntcol)
    elif entrnames[it][0] == 'button':
        submit = ttk.Button(mainframe, text= entrnames[it][1], font=16)
        submit.grid(row = cntrow, column = cntcol)
    else:
        listout[-1].set(entrnames[it][1]) # set the default option
        popupMenu = OptionMenu(mainframe, listout[-1], *choices)
        Label(mainframe, text= entrnames[it][2]).grid(row = cntrow, column = cntcol)
        popupMenu.grid(row = cntrow+1, column = cntcol)            
    if cntrow >= rowN-2:
        cntrow = 1
        cntcol = cntcol+1
    else:
        cntrow = cntrow+2
root.mainloop()

#%%


# Create a list of Tkinter variables



tkvar = StringVar(root)
#tkvar = TextEntry(root)

# 1) text box, Name of Course
Label(mainframe, text="Name of Course").grid(row = 1, column = 1)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 2, column = 1)

# 2) text box, Starting date Course
Label(mainframe, text="Starting date").grid(row = 3, column = 1)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 4, column = 1)

# 3) text box, Amount of weeks
Label(mainframe, text="Amount of weeks").grid(row = 5, column = 1)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 6, column = 1)

# 4) text box, Time of class
Label(mainframe, text="Time of class").grid(row = 7, column = 1)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 8, column = 1)

# 5) text box, Location of class
Label(mainframe, text="Location of class").grid(row = 9, column = 1)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 10, column = 1)

# 6) text box, Amount of Teachers
Label(mainframe, text="Amount of teachers").grid(row = 9, column = 1)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 10, column = 1)

# 7) Dictionary with options
tkvar.set('Solo') # set the default option
popupMenu = OptionMenu(mainframe, tkvar, *choices)
Label(mainframe, text="What type of class is this").grid(row = 1, column = 2)
popupMenu.grid(row = 2, column = 2)

# 8) text box, Max people/couples
Label(mainframe, text="Max amount of people/couples").grid(row = 3, column = 2)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 4, column = 2)

# 9) text box, Max extras
Label(mainframe, text="Max amount of extras").grid(row = 5, column = 2)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 6, column = 2)

# 10) text box, Excel file ID
Label(mainframe, text="google sheet file ID").grid(row = 7, column = 2)
entry = ttk.Entry(mainframe, font=16)
entry.grid(row = 8, column = 2)

# 11) ok button
submit = ttk.Button(mainframe, text='OK', font=16)
submit.grid(row = 10, column = 2)

# on change dropdown value
def change_dropdown(*args):
    print( tkvar.get() )

# link function to change dropdown
tkvar.trace('w', change_dropdown)

root.mainloop()