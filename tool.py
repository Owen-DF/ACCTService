#Necessarry Libraries
from tkinter import *
import os
import pandas as pd
import openpyxl
import warnings

#functions
def submit():

    directory = Dinput.get('1.0', END).strip()
    print(directory)
    
    items = IInput.get('1.0', END).strip()
    itemList = items.split(", ")
    print(itemList)



    return




#GUI set up

root = Tk()
root.title("ACCT Tool")
root.geometry('400x200')
root.resizable(0,0)
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=3)




#directory input
DLabel = Label(text = "Enter top directory:")
Dinput = Text(root, height = "1", width = 30)
DLabel.grid(column = 0, row =0, sticky=W, padx=5, pady = 5)
Dinput.grid(column = 1,row = 0, sticky=W, padx=5, pady=5)


#item input
ILabel = Label(text = "Enter items to search:")
IInput = Text(root, height = "1", width = 30) 
ILabel.grid(column = 0, row = 1, sticky = W, padx = 5, pady = 5)
IInput.grid(column = 1, row = 1, sticky = W, padx = 5, pady = 5)
IInput.insert('1.0', "seperate items with comas")


submit = Button(text="Enter", command = submit)
submit.grid(column = 0, row = 2, padx=5, pady=5)


root.mainloop()
