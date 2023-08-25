import pandas as pd

# Load the Excel file
xls = pd.ExcelFile('new.xlsx')

# Get the sheet names
sheet_names = xls.sheet_names

# Read only the first sheet into a dataframe
df = pd.read_excel('new.xlsx', sheet_name=sheet_names[0])


df.to_excel('new.xlsx', sheet_name=sheet_names[0], index=False)
