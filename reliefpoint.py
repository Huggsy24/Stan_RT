import pandas as pd
import numpy as np
import math

# Function to process each table
def process_table(df):
    # Initialize variables for service ID and pattern
    service_id = "Service ID not found"
    pattern = "Pattern not found"
    service_id_found, pattern_found = False, False  # Flags to indicate if they are found
    max_distance = 10  # Maximum distance between 'Service ID:' and 'Pattern:'

    # Iterate through the DataFrame to find 'Service ID:' and 'Pattern:'
    for index, row in df.iterrows():
        if "Service ID:" in str(row[0]):
            service_id = row[1] if pd.notnull(row[1]) else service_id
            service_id_found = True
        if "Pattern:" in str(row[0]):
            pattern = row[1] if pd.notnull(row[1]) else pattern
            pattern_found = True
        if service_id_found and pattern_found:
            break  # Found both, break the loop
        if (service_id_found or pattern_found) and (index > max_distance):
            # If only one is found and max_distance is exceeded, reset flags and values
            service_id, pattern = "Service ID not found", "Pattern not found"
            service_id_found, pattern_found = False, False

    # Skip processing if DataFrame is empty or only contains header information
    if df.shape[0] <= 2:
        return pd.DataFrame()

    # Add service ID and pattern to the top of the DataFrame
    header_rows = [['Service ID', service_id], ['Pattern', pattern]]
    header_df = pd.DataFrame(header_rows, columns=df.columns[:2])  # Adjust columns to match df structure

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
            average_running_times = [math.ceil(x) if not np.isnan(x) else np.nan for x in running_times_df.mean(skipna=True)]
        else:
            average_running_times = [np.nan] * (block_df.shape[1] - 1)  # exclude the time interval column

        return [hour_block] + average_running_times

    updated_rows_corrected = []
    for hour_block in standard_time_intervals:
        updated_row = calculate_average_for_hour_block_corrected(hour_block, df.iloc[2:])  # Adjust to skip header rows
        updated_rows_corrected.append(updated_row)

    # Create the DataFrame with updated values
    standardized_and_updated_df = pd.DataFrame(updated_rows_corrected, columns=['Time Interval'] + list(df.columns[1:]))

    # Concatenate header_df and standardized_and_updated_df to add header rows without adding an empty DataFrame
    final_df = pd.concat([header_df, standardized_and_updated_df], ignore_index=True)

    return final_df

# Load the Excel file and the specific sheet
file_path = '/Users/hugo.cooke/Desktop/test/2copy.xlsx'  # Update this to your actual file path
df = pd.read_excel(file_path, sheet_name='Travel times Outbound')  # Make sure the sheet name matches

# Initialize list to hold processed tables
processed_tables = []

# Initialize flag to determine if a new table is being processed
processing_table = False
current_table = []

# Loop through rows of the DataFrame
for index, row in df.iterrows():
    if "Service ID:" in str(row[0]) or "Pattern:" in str(row[0]):
        # If a new table is detected, process the previous one
        if processing_table:
            processed_table = process_table(pd.DataFrame(current_table, columns=df.columns))
            if not processed_table.empty:  # Check if the processed table is not empty
                processed_tables.append(processed_table)
            current_table = []  # Reset the current table
        processing_table = True

    if processing_table:
        current_table.append(row)

# Process the last table
if current_table:
    processed_table = process_table(pd.DataFrame(current_table, columns=df.columns))
    if not processed_table.empty:  # Check if the processed table is not empty
        processed_tables.append(processed_table)

# Concatenate all processed tables vertically
final_df = pd.concat(processed_tables, ignore_index=True)

# Save the DataFrame to a new Excel file
output_file_path = '/Users/hugo.cooke/Desktop/test/cleaned.xlsx'  # Update this to the desired output file path
final_df.to_excel(output_file_path, index=False)

print(f"Data has been cleaned and saved to {output_file_path}")
