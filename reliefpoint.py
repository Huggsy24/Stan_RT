import pandas as pd
import numpy as np
import math

# Load the Excel file and the specific sheet
file_path = '/Users/hugo.cooke/Desktop/test/2copy.xlsx'  # Update this to your actual file path
df = pd.read_excel(file_path, sheet_name='Travel times Outbound')  # Make sure the sheet name matches

# Define the standard time intervals
standard_time_intervals = ['0000-0059', '0100-0159', '0200-0259', '0300-0359', '0400-0459', '0500-0559', '0600-0659', '0700-0759', '0800-0859', '0900-0959', '1000-1059', '1100-1159', '1200-1259', '1300-1359', '1400-1459', '1500-1559', '1600-1659', '1700-1759', '1800-1859', '1900-1959', '2000-2059', '2100-2159', '2200-2259', '2300-2359']

# Function to calculate averages and handle only numeric columns
def calculate_average_for_hour_block_corrected(hour_block, block_df):
    hour_block_start, hour_block_end = map(lambda x: int(x), hour_block.split('-'))
    overlapping_rows = []
    for _, row in block_df.iterrows():
        if pd.isnull(row.iloc[0]) or '-' not in str(row.iloc[0]):
            continue
        row_interval_start, row_interval_end = map(lambda x: int(x), str(row.iloc[0]).split('-'))
        if (row_interval_start < hour_block_end) and (hour_block_start < row_interval_end):
            overlapping_rows.append(row[1:])  # Exclude the first column which contains the time interval
    
    if overlapping_rows:
        running_times_df = pd.DataFrame(overlapping_rows)
        # Calculate the average for each column, ignoring NaNs, and round up
        average_running_times = [math.ceil(x) if not np.isnan(x) else np.nan for x in running_times_df.mean(axis=0, skipna=True)]
    else:
        average_running_times = [np.nan] * (block_df.shape[1] - 1)  # exclude the time interval column
    
    return [hour_block] + average_running_times

# Process a subset or block of data from 'df' as an example
following_rows_df = df.iloc[:10, :]  # Example, adjust this to select the correct block of data

updated_rows_corrected = []
for hour_block in standard_time_intervals:
    updated_row = calculate_average_for_hour_block_corrected(hour_block, following_rows_df)
    updated_rows_corrected.append(updated_row)

# Create the DataFrame with updated values
standardized_and_updated_first_block_df_corrected = pd.DataFrame(updated_rows_corrected, columns=['Time Interval'] + following_rows_df.columns[1:].tolist())

# Save the DataFrame to a new Excel file
output_file_path = '/Users/hugo.cooke/Desktop/test/cleaned.xlsx'  # Update this to the desired output file path
standardized_and_updated_first_block_df_corrected.to_excel(output_file_path, index=False)

print(f"Data has been cleaned and saved to {output_file_path}")
