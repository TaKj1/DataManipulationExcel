import pandas as pd

# Load the Excel file
xls = pd.ExcelFile('/home/genadi/Desktop/all.xlsx', engine='openpyxl')

# Function to identify empty header cells in a dataframe
def identify_empty_headers(df):
    empty_headers = []
    # We only need to check the first row for headers
    for idx, value in enumerate(df.columns):
        if pd.isna(value) or "Unnamed" in str(value):
            empty_headers.append(idx)
    return empty_headers


for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl')
    empty_header_indices = identify_empty_headers(df)
    if empty_header_indices:
        for idx in empty_header_indices:
            # Report sheet name, column index (as Excel letter), and row number
            col_letter = chr(65 + idx)  # Convert 0-based index to Excel-style column letter, assuming it's within Z.
            print(f"In sheet '{sheet_name}', empty header in column '{col_letter}', row 1.")

