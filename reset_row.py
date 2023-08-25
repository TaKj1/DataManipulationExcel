def reset_row(df, row_num, start_col_index=2):
    
    for col_index in range(start_col_index, df.shape[1]):
        try:
            df.iat[row_num-2, col_index] = 0
        except Exception as e:
            print(f"Failed to reset cell at row {row_num}, column index {col_index}. Error: {e}")
    
    print(f"\nData types in row {row_num} after attempted reset:")