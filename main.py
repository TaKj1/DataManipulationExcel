import pandas as pd
import string
from move_row import get_last_filled_value_before_empty
from reset_row import reset_row

# Function to convert column index to Excel-style column letter
def col_index_to_letter(index):
    letters = string.ascii_uppercase
    if index < 26:
        return letters[index]
    else:
        return letters[index // 26 - 1] + letters[index % 26]

filename = '/home/genadi/Desktop/new.xlsx'
xls = pd.ExcelFile(filename)          # Load the Excel file

sheet_names = xls.sheet_names                   # Get the sheet names

# Read only the first sheet into a dataframe
df = pd.read_excel(xls, sheet_name=sheet_names[0])

# Prompt user to choose action
action = input("Choose an action (move_row or reset_row): ").strip().lower()

# Validate user input
while action not in ["move_row", "reset_row"]:
    action = input("Invalid choice. Please choose either 'move_row' or 'reset_row': ").strip().lower()

# Accept row numbers from the user
row_nums = input("Enter row numbers, separated by commas (e.g., 10,11,12): ")
row_nums = [int(row.strip()) for row in row_nums.split(",")]

for row_num in row_nums:
    if action == "reset_row":
        reset_row(df, row_num)
        print(f"Row {row_num} after reset:")                        
        print(df.iloc[row_num-2])
    elif action == "move_row":
        get_last_filled_value_before_empty(df)
        print(f"Row {row_num} after move:")
        # Note: If move_row moves the row to the end, 
        # you might want to print the last row here instead.
        ####print('-' * 40)

##df.to_excel(filename, sheet_name=sheet_names[0], index=False)  # Save changes back to the Excel file
