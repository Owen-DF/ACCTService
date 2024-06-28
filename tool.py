#Necessarry Libraries

from tkinter import *
import ttkbootstrap as ttk
import os
import pandas as pd
import openpyxl
import warnings
import glob
import time
 
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)
 

#file paths for reference, change here for future file reconstructing
YEARLY_COSTING = r"C:\Users\ofleischer\OneDrive - MetroTek Electrical\Coding\SampleFiles\Yearly Costing.xlsx"  #have
BLANK_CELL = 'G101'
OPENPOS = r'Z:\Accounting\13 Python Reports\Open POs.xlsx'   #need
JOBHOURS = r'Z:\Accounting\13 Python Reports\Job Hours.xlsx' #need


#functions
def submit():

    directory = DInput.get().strip()
    bar['value'] = 0
    items = IInput.get().strip()
    itemList = filterProjects(items)
    bar.pack()
    fileNav(directory, itemList)
    return



def filterProjects(items):

    itemList = [item.strip() for item in items.split(",") if item.strip()]
    itemList = [item for item in itemList if len(item) == 4]
    uniqueList = remove_duplicates(itemList)
    return uniqueList

def remove_duplicates(item_list):
    seen = set()
    unique_list = []
    for item in item_list:
        if item not in seen:
            unique_list.append(item)
            seen.add(item)
    return unique_list


def fileNav(filePath, itemList):
    matching_folders = []
    for item in itemList:
        # Create a pattern to match folders starting with the project code
        pattern = os.path.join(filePath, f"{item}*")
        #print(f"Searching with pattern: {pattern}")
        # Use glob to find matching directories
        for folder in glob.glob(pattern):
            if os.path.isdir(folder):
                folder = os.path.join(folder, '03 Estimate & Proposal')
                if os.path.isdir(folder):
                    matching_folders.append(folder)
                   # print(f"Found matching folder: {folder}")
    
    pullExcel(matching_folders, itemList)
    return



def pullExcel(matchingFolders, itemList):
    ExcelFilePaths = []
    if not matchingFolders:
      #  print("No folders to process")
        return
    try:
        stepValue = 100 / len(matchingFolders)
    except ZeroDivisionError:
      #  print("Cannot divide by zero, most likely no input in form")
        return
    for folder in matchingFolders:
        excelFiles = []
        pattern = os.path.join(folder, '*.xlsm')  #these excel files are not xlsx, whoever decided that tbere are different types of excel files deserve something bad, i am cranky
       # print(f"Searching for Excel files with pattern: {pattern}")       
        for file in glob.glob(pattern):
            excelFiles.append(file)
           # print(f"Found Excel File: {file}")
        if excelFiles:
            if len(excelFiles) > 1:
                mostRecent = max(excelFiles, key=os.path.getatime)
                ExcelFilePaths.append(mostRecent)
            else:
                ExcelFilePaths.append(excelFiles[0])
        else:
            return
        bar['value'] += stepValue
        app.update_idletasks()

    PreProcessExcelFiles(ExcelFilePaths, itemList)


def PreProcessExcelFiles(ExcelList, itemList):
    df = pd.concat([pd.read_excel(YEARLY_COSTING, sheet_name=sheet_name) for sheet_name in ['2024','2023', '2022', '2021', '2020', '2019', '2018']])
    df.fillna('', inplace=True)
    

    for i, file_path in enumerate(ExcelList):
        searchTerm = itemList[i]
        processExcel(file_path, searchTerm, df)

def processExcel(filePath, searchTerm, df):
    items = {
        '01.01': 'G9', '01.02': 'G10', '02.01': 'G21', '02.02': 'G22', '02.03': 'G23',
        '03.01': 'G37', '03.02': 'G38', '03.03': 'G39', '04.01': 'G48', '04.02': 'G49',
        '04.03': 'G50', '04.04': 'G51', '05.01': 'G63', '05.02': 'G64', '05.03': 'G65',
        '05.04': 'G66', '06.01': 'G78', '06.02': 'G79', '06.03': 'G80', '06.04': 'G81',
        '06.05': 'G82', '07.0': 'G90', '08.01': 'G94', '08.02': 'G95', '08.03': 'G96','08.04': 'G97'
    }
    
    workbook = openpyxl.load_workbook(filePath, read_only=False, keep_vba = True)
    worksheet = workbook['JC']

    for item, cell in items.items():
        result = df.loc[(df['Name'].astype(str).str.contains(searchTerm)) & (df['Item'].astype(str).str.contains(item)), 'Amount'].sum()
        worksheet.save(filePath)
    workbook.save(filePath)


    result = df.loc[(df['Name'].astype(str).str.contains(searchTerm))&(df['Item'] == ''), 'Amount'].sum()
    cell = BLANK_CELL
    worksheet[cell] = result
    workbook.save(filePath)

    df_open_pos = pd.read_excel(OPENPOS)
    items = {
        '01.01': 'I9', '01.02': 'I10', '02.01': 'I21', '02.02': 'I22', '02.03': 'I23',
        '03.01': 'I37', '03.02': 'I38', '03.03': 'I39', '04.01': 'I48', '04.02': 'I49',
        '04.03': 'I50', '04.04': 'I51', '05.01': 'I63', '05.02': 'I64', '05.03': 'I65',
        '05.04': 'I66', '06.01': 'I78', '06.02': 'I79', '06.03': 'I80', '06.04': 'I81',
        '06.05': 'I82', '07.0': 'I90', '08.01': 'I94', '08.02': 'I95', '08.03': 'I96','08.04': 'I97'
    }

    workbook = openpyxl.load_workbook(filePath, read_only = False, keep_vba = True)
    worksheet = workbook['JC']

    for item, cell in items.items():
        result = df_open_pos.loc[(df_open_pos['Name'].astype(str).str.contains(searchTerm)) & (df_open_pos['Item'].astype(str).str.contains(item)), 'Open Balance'].sum()
        worksheet[cell] = result
    workbook.save(filePath)

    dfJobHours = pd.read_excel(JOBHOURS)
    totalAmount = dfJobHours.loc[dfJobHours['Job Description'].astyhpe(str).str.contains(searchTerm), 'Hours'].sum()

    cell = 'G87'
    worksheet[cell] = totalAmount
    workbook.save(filePath)
    
    return


#GUI set up


app = ttk.Window(themename="superhero")
app.title("EasyAccounting")


app.geometry("600x500")

label = ttk.Label(app, text="EasyAccounting") # Creates a label
label.pack(pady=30) # Pack label in window 
label.config(font=("Times New Roman", 20, "bold")) # Increase font size & make it bold

path_frame = ttk.Frame(app) # Creates frame
path_frame.pack(pady=15, padx=10, fill="x") # Pack frame in app
ttk.Label(path_frame, text="Full Directory").pack(side="left", padx=5) # Create & pack label
DInput = ttk.Entry(path_frame)
DInput.pack(side="left", fill="x", expand=True, padx=5) # Create & pack entry widget

codes_frame = ttk.Frame(app) # Creates frame
codes_frame.pack(pady=15, padx=10, fill="x") # Pack frame in app
ttk.Label(codes_frame, text="Project Codes").pack(side="left", padx=5) # Create & pack label
IInput = ttk.Entry(codes_frame)
IInput.pack(side="left", fill="x", expand=True, padx=5) # Create & pack entry widget



button_frame = ttk.Frame(app) # Create a frame for buttons
button_frame.pack(pady=50, padx=10, fill="x") # Pack frame in app
ttk.Button(button_frame, text="Submit", command = submit, bootstyle = "success").pack(side="left", padx=10) # Create & pack button

progress_frame = ttk.Frame(app)
progress_frame.pack(pady=25, padx = 10, fill="x")
bar = ttk.Progressbar(progress_frame, length = 400, style='success.Striped.Horizontal.TProgressbar')
# bar.pack()


app.mainloop()
