import pandas as pd

# Load the Excel file
xls = pd.ExcelFile('/home/genadi/Desktop/all.xlsx', engine='openpyxl')

# Extract immutable rows from the first sheet
immutable_rows = pd.read_excel(xls, sheet_name=xls.sheet_names[0], engine='openpyxl', nrows=4)

# Start with the data from the 8th row of the first sheet to initiate the consolidated dataframe
df_main = pd.read_excel(xls, sheet_name=xls.sheet_names[0], engine='openpyxl', skiprows=6)

# Keep the first four columns intact
df_main = df_main[df_main.columns[:4]]

# Loop through all the sheets
for sheet_name in xls.sheet_names[1:]:  # Starting from the second sheet
    # Extract the first 4 rows from the current sheet
    current_immutable_rows = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl', nrows=4)
    
    # Check if these rows exist in the immutable_rows dataframe
    if not current_immutable_rows.equals(immutable_rows):
        immutable_rows = pd.concat([immutable_rows, current_immutable_rows], ignore_index=True)

    # Process the rest of the rows
    df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl', skiprows=6)
    
    # Get the last 8 columns from the current sheet
    columns_to_append = df[df.columns[-8:]]
    
    # Rename the last 8 columns with the month prefix
    columns_to_append.columns = [f"{sheet_name}_{col}" for col in columns_to_append.columns]
    
    # Append these columns to the main dataframe
    df_main = pd.concat([df_main, columns_to_append], axis=1)

# Replace NaN values with 0
df_main.fillna(0, inplace=True)

# Save the consolidated data to a new Excel sheet
with pd.ExcelWriter('consolidated_data.xlsx', engine='openpyxl') as writer:
    # Write the immutable rows first
    immutable_rows.to_excel(writer, sheet_name='All Data', index=False, startrow=0, header=False)
    # Write the consolidated data, starting after the immutable rows
    df_main.to_excel(writer, sheet_name='All Data', index=False, startrow=len(immutable_rows))
