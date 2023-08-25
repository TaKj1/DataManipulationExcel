import pandas as pd

# Load your Excel file
xls = pd.ExcelFile('/home/genadi/Desktop/all.xlsx', engine='openpyxl')

# Loop through each sheet and identify empty columns
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl')
    
   
    empty_columns = df.columns[df.isnull().all()].tolist()

    if empty_columns:
        print(f"In sheet '{sheet_name}', the empty columns are: {', '.join(empty_columns)}")
    else:
        print(f"In sheet '{sheet_name}', there are no empty columns.")
