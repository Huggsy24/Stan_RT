import pandas as pd
import numpy as np
import math

# Function to process each table
def process_table(df):
    service_id = "Service ID not found"
    pattern = "Pattern not found"
    service_id_found, pattern_found = False, False
    max_distance = 10

    for index, row in df.iterrows():
        if "Service ID:" in str(row[0]):
            service_id = row[1] if pd.notnull(row[1]) else service_id
            service_id_found = True
        if "Pattern:" in str(row[0]):
            pattern = row[1] if pd.notnull(row[1]) else pattern
            pattern_found = True
        if service_id_found and pattern_found:
            break
        if (service_id_found or pattern_found) and (index > max_distance):
            service_id, pattern = "Service ID not found", "Pattern not found"
            service_id_found, pattern_found = False, False

    if df.shape[0] <= 2:
        return pd.DataFrame()

    header_rows = [['Service ID', service_id], ['Pattern', pattern]]
    header_df = pd.DataFrame(header_rows, columns=['Info', 'Value'])

    standard_time_intervals = ['0000-0059', '0100-0159', '0200-0259', '0300-0359', '0400-0459', '0500-0559', '0600-0659', '0700-0759', '0800-0859', '0900-0959', '1000-1059', '1100-1159', '1200-1259', '1300-1359', '1400-1459', '1500-1559', '1600-1659', '1700-1759', '1800-1859', '1900-1959', '2000-2059', '2100-2159', '2200-2259', '2300-2359']

    def calculate_average_for_hour_block_corrected(hour_block, block_df):
        hour_block_start, hour_block_end = map(lambda x: int(x), hour_block.split('-'))
        overlapping_rows = []
        for _, row in block_df.iterrows():
            if pd.isnull(row.iloc[0]) or '-' not in str(row.iloc[0]):
                continue
            row_interval_start, row_interval_end = map(lambda x: int(x), str(row.iloc[0]).split('-'))
            if (row_interval_start < hour_block_end) and (hour_block_start < row_interval_end):
                overlapping_rows.append(row[1:])

        if overlapping_rows:
            running_times_df = pd.DataFrame(overlapping_rows)
            average_running_times = [math.ceil(x) if not np.isnan(x) else np.nan for x in running_times_df.mean(skipna=True)]
        else:
            average_running_times = [np.nan] * (block_df.shape[1] - 1)

        return [hour_block] + average_running_times

    updated_rows_corrected = []
    for hour_block in standard_time_intervals:
        updated_row = calculate_average_for_hour_block_corrected(hour_block, df.iloc[2:])  # Adjust to skip header rows
        updated_rows_corrected.append(updated_row)

    standardized_and_updated_df = pd.DataFrame(updated_rows_corrected, columns=['Time Interval'] + list(df.columns[1:]))

    # Concatenation of header_df and standardized_and_updated_df without setting custom column names
    final_df = pd.concat([header_df, standardized_and_updated_df], ignore_index=True)
    
    # Removed the explicit setting of column names
    
    return final_df

# Define a function to process a given sheet and return the processed DataFrame
def process_sheet(sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    processed_tables = []
    processing_table = False
    current_table = []

    for index, row in df.iterrows():
        if "Service ID:" in str(row[0]) or "Pattern:" in str(row[0]):
            if processing_table:
                processed_table = process_table(pd.DataFrame(current_table, columns=df.columns))
                if not processed_table.empty:
                    processed_tables.append(processed_table)
                current_table = []
            processing_table = True

        if processing_table:
            current_table.append(row)

    if current_table:
        processed_table = process_table(pd.DataFrame(current_table, columns=df.columns))
        if not processed_table.empty:
            processed_tables.append(processed_table)

    final_df = pd.concat(processed_tables, ignore_index=True)
    return final_df

# Specify the Excel file path
file_path = '/Users/hugo.cooke/Desktop/test/2copy.xlsx'

# Process both sheets
outbound_df = process_sheet('Travel times Outbound')
inbound_df = process_sheet('Travel times Inbound')

# Save the DataFrames to separate sheets within the same Excel file
output_file_path = '/Users/hugo.cooke/Desktop/test/cleaned.xlsx'
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    outbound_df.to_excel(writer, sheet_name='Outbound', index=False)
    inbound_df.to_excel(writer, sheet_name='Inbound', index=False)

print(f"Data has been cleaned and saved to {output_file_path}")
