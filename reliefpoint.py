import pandas as pd
import numpy as np
import math
import glob
import os
import zipfile


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

    standard_time_intervals = ['0000-0559', '0600-0659', '0700-0729', '0730-0744', '0745-0759', '0800-0814', '0815-0829', '0830-0844', '0845-0859', '0900-0929', '0930-1159', '1200-1359', '1400-1529', '1530-1629', '1630-1659', '1700-1729', '1730-1759', '1700-1759', '1800-1829', '1830-1929', '1930-2059', '2100-3559']

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
def process_sheet(sheet_name, file_path):  # Add file_path as a second parameter
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')  # Use the file_path parameter here
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

def process_all_sheets(file_path):
    xl = pd.ExcelFile(file_path, engine='openpyxl')  # Use 'openpyxl' engine for reading .xlsx files
    sheet_names = [name for name in xl.sheet_names if "travel times " in name.lower()]
    outbound_dfs = []
    inbound_dfs = []
    
    for sheet_name in sheet_names:
        try:
            df = process_sheet(sheet_name, file_path)
            if not df.empty:  # Check if the dataframe is empty
                if "outbound" in sheet_name.lower():
                    outbound_dfs.append(df)
                elif "inbound" in sheet_name.lower():
                    inbound_dfs.append(df)
        except Exception as e:
            print(f"Skipping sheet '{sheet_name}' due to error: {e}")
            continue
    
    # Initialize empty dataframes if no sheets were processed successfully
    if not outbound_dfs:
        outbound_df = pd.DataFrame()
    else:
        outbound_df = pd.concat(outbound_dfs, ignore_index=True)
    
    if not inbound_dfs:
        inbound_df = pd.DataFrame()
    else:
        inbound_df = pd.concat(inbound_dfs, ignore_index=True)
    
    return outbound_df, inbound_df

# Loop through all .xlsx files in the specified folder
folder_path = '/Users/hugo.cooke/Desktop/test'  # Adjust the folder path as necessary
for file_path in glob.glob(os.path.join(folder_path, '*.xlsx')):
    try:
        outbound_df, inbound_df = process_all_sheets(file_path)

        # Generate a new output filename by appending '_cleaned' before the file extension
        base_filename = os.path.basename(file_path)
        name_part, extension_part = os.path.splitext(base_filename)
        output_filename = f"{name_part}_cleaned{extension_part}"
        output_file_path = os.path.join(folder_path, output_filename)

        # Save the processed data to a new file
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            outbound_df.to_excel(writer, sheet_name='Outbound', index=False)
            inbound_df.to_excel(writer, sheet_name='Inbound', index=False)

        print(f"Data has been cleaned and saved to {output_file_path}")
    except zipfile.BadZipFile:
        print(f"Skipping file due to error (not a zip file): {file_path}")
    except Exception as e:
        print(f"Skipping file due to unexpected error: {file_path}, Error: {e}")
