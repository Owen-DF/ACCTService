import pandas as pd
import openpyxl
import warnings
 
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=FutureWarning)
 
def process_excel_file(file_path, search_term, df):
    items = {
        '01.01': 'G9', '01.02': 'G10', '02.01': 'G21', '02.02': 'G22', '02.03': 'G23',
        '03.01': 'G37', '03.02': 'G38', '03.03': 'G39', '04.01': 'G48', '04.02': 'G49',
        '04.03': 'G50', '04.04': 'G51', '05.01': 'G63', '05.02': 'G64', '05.03': 'G65',
        '05.04': 'G66', '06.01': 'G78', '06.02': 'G79', '06.03': 'G80', '06.04': 'G81',
        '06.05': 'G82', '07.0': 'G90', '08.01': 'G94', '08.02': 'G95', '08.03': 'G96','08.04': 'G97'
    }
 
    workbook = openpyxl.load_workbook(file_path, read_only=False, keep_vba=True)
    worksheet = workbook['JC']
 
    for item, cell in items.items():
        result = df.loc[(df['Name'].astype(str).str.contains(search_term)) & (df['Item'].astype(str).str.contains(item)), 'Amount'].sum()
        worksheet[cell] = result
    workbook.save(file_path)
 
    result = df.loc[(df['Name'].astype(str).str.contains(search_term)) & (df['Item'] == ''), 'Amount'].sum()
    cell = 'G101'
    worksheet[cell] = result
    workbook.save(file_path)
 
    df_open_pos = pd.read_excel(r'Z:\Accounting\13 Python Reports\Open POs.xlsx')
 
    items = {
        '01.01': 'I9', '01.02': 'I10', '02.01': 'I21', '02.02': 'I22', '02.03': 'I23',
        '03.01': 'I37', '03.02': 'I38', '03.03': 'I39', '04.01': 'I48', '04.02': 'I49',
        '04.03': 'I50', '04.04': 'I51', '05.01': 'I63', '05.02': 'I64', '05.03': 'I65',
        '05.04': 'I66', '06.01': 'I78', '06.02': 'I79', '06.03': 'I80', '06.04': 'I81',
        '06.05': 'I82', '07.0': 'I90', '08.01': 'I94', '08.02': 'I95', '08.03': 'I96','08.04': 'I97'
    }
 
    workbook = openpyxl.load_workbook(file_path, read_only=False, keep_vba=True)
    worksheet = workbook['JC']
 
    for item, cell in items.items():
        result = df_open_pos.loc[(df_open_pos['Name'].astype(str).str.contains(search_term)) & (df_open_pos['Item'].astype(str).str.contains(item)), 'Open Balance'].sum()
        worksheet[cell] = result
    workbook.save(file_path)
 
    df_job_hours = pd.read_excel(r'Z:\Accounting\13 Python Reports\Job Hours.xlsx')
    total_amount = df_job_hours.loc[df_job_hours['Job Description'].astype(str).str.contains(search_term), 'Hours'].sum()
 
    cell = 'G87'
    worksheet[cell] = total_amount
    workbook.save(file_path)
 




file_paths = [
    r'Z:\Projects\02 ACTIVE PROJECTS\9195-Maddox-Transformer Repair-Bronx NY\03 Estimate & Proposal\MTES24-009-Maddox-Transformer Repair-Bronx NY Rev2.xlsm',
]
 
search_terms = [
    '9195',
]

df = pd.concat([pd.read_excel(r'Z:\Accounting\13 Python Reports\Yearly Costing.xlsx', sheet_name=sheet_name) for sheet_name in ['2024','2023', '2022', '2021', '2020', '2019', '2018']])
df.fillna('', inplace=True)
 
for i, file_path in enumerate(file_paths):
    search_term = search_terms[i]
    process_excel_file(file_path, search_term, df)
    print(f"{search_term} complete.")