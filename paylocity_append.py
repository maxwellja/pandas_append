#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import openpyxl as op
import re


# In[4]:


file_path = 'your_path_to.xlsx'
data = pd.ExcelFile(file_path)
all_sheets = data.sheet_names

#Retrieves 'template' table in the excel file, used to test column headers.
xls = op.load_workbook(file_path)
template_name = 'TEMPLATE'
tbl_r = xls[template_name].tables.items()[0][1]

#Disregards sheets outside of naming convention & stores them for exception handling later.
wk_sheets = []
ignored = []
excluded = [template_name, "Combined"]
regex = r"(\d|1[0-2])-2\d-(FP|INC|ALL)" #Regex for sheets matching format [month]-[year]-[subsidiary].
for sheet in all_sheets:
    if re.match(regex, sheet):
        wk_sheets.append(sheet)
    elif sheet not in excluded:
        ignored.append(sheet)

#Returns data and subsidiary from sheet names.
def parse_sheet_name(sheet_name):
    parts = sheet_name.split('-')
    date_str = f"20{parts[1]}-{parts[0]}-01"  # Assuming the format is Month-Year-Subsidiary
    return pd.to_datetime(date_str), parts[2]


# In[5]:


# Create a DataFrame to hold combined data.

#To validate all sheet headers match.
headers = pd.read_excel(data, sheet_name=template_name, skiprows=int(tbl_r[1])-1, nrows=1, usecols=tbl_r[0]+":"+tbl_r[3])
status = 'Employee Status - Current'
hire = 'Hire Date - Current'
rehire = 'Rehire Date - Current'
e_id = 'Employee ID'

# Lists to store iterated dfs; parsing errors.
df_sheets = []
parsing_errors = []
bad_headers = []

# Process each sheet:
for sheet in wk_sheets:
    df = pd.read_excel(data, sheet_name=sheet)
    test_cols = list(df.columns.difference(headers.columns))
    date, subsidiary = parse_sheet_name(sheet)
    
    df['Report Date'] = date  # Add sheet date as a new column
    
    if test_cols != []:
        print(df.columns, test_cols)
        bad_headers.append(sheet)
        continue
    elif subsidiary == 'FP':
        df[e_id] = df[e_id].map(lambda x: f"FP - {x}") # Add 'FP - ' to each employee ID
        df_sheets.append(df)
    elif subsidiary in {'INC', 'ALL'}:
        df_sheets.append(df)
    else:
        parsing_errors.append(f"Error parsing sheet: {sheet}")

#Concatenate dataframes together.
if bad_headers != []:
    print("There are sheets with column names inconsistent with the template. Check the following sheets:")
    for sheet in bad_headers:
        print(sheet)
else:
    appended = pd.concat(df_sheets)
    #Notifications:
    print("Complete")
    if parsing_errors != []:
        print("The following pages could not be parsed, so was/were ignored:")
        for error in parsing_errors:
            print(error)
    if ignored != []:
        print("The following pages didn't match the sheet naming convention, so was/were ignored:")
        for name in ignored:
            print(name)

appended.reset_index(drop=True, inplace=True)

#Cleaning Steps
appended[status] = appended[status].str[:1] #Trunkcate first letter of Employee Status
appended.loc[appended[rehire].notna(), [hire]] = appended[rehire] #Change hire date to rehire date, if appl.
appended = appended.drop(rehire, axis = 1) #Delete rehire column


# In[6]:


# Save the combined data to a new Excel file
file_title = "clean_appeneded_data"
destination = f"path_{file_title}.xlsx"
appended.to_excel(destination, index=False)
print("Complete")

