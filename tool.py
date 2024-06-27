#Necessarry Libraries

from tkinter import *
import os
import pandas as pd
import openpyxl
import warnings
import glob
#functions
def submit():

    directory = Dinput.get('1.0', END).strip()
    
    items = IInput.get('1.0', END).strip()
    itemList = items.split(", ")

    fileNav(directory, itemList)
    return


def fileNav(filePath, itemList):
    matching_folders = []
    for item in itemList:
        # Create a pattern to match folders starting with the project code
        pattern = os.path.join(filePath, f"{item}*")
        print(f"Searching with pattern: {pattern}")
        # Use glob to find matching directories
        for folder in glob.glob(pattern):
            if os.path.isdir(folder):
                folder = os.path.join(folder, '03 Estimate & Proposal')
                if os.path.isdir(folder):
                    matching_folders.append(folder)
                    print(f"Found matching folder: {folder}")
    pullExcel(matching_folders)
    return
            
def pullExcel(matchingFolders):
    for folder in matchingFolders:
        pattern = os.path.join(folder, '*.xlsx')
        print(f"Searching for Excel files with pattern: {pattern}")
        for file in glob.glob(pattern):
            print(f"Found Excel File: {file}")

            


    





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
