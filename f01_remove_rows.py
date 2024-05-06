import re
import pandas as pd
import numpy as np
from pathlib import Path


def read_input_file_type(input_file_path, input_file_type, sheet_name=None):
    if input_file_type == '.xlsx':
        return pd.read_excel(input_file_path, sheet_name=sheet_name, header=None)
    elif input_file_type == '.csv':
        return pd.read_csv(input_file_path, header=None)
    elif input_file_type == '.txt':
        return pd.read_csv(input_file_path, delimiter='\t', header=None)
    else:
        raise ValueError(f"Unsupported file type: {input_file_type}")
    
        
def apply_removal_criteria(df, criteria):
    
    column = criteria['applied_column']
    if column not in df.columns:
        print(f"The column '{column}' does not exist in the dataframe. Skipping the removal criteria.")
        return df
    
    # column = criteria['applied_column']
    operation = criteria['criteria_to_remove_row']
    value = criteria['criteria_value']
    
    if operation == '=':
        # Remove rows where the column value equals the criteria value
        df = df[df[column].astype(str).str.strip() != value]
    elif operation == '> =':
        # Remove rows where the column value is less than the criteria value
        df = df[df[column] < float(value)]
    elif operation == '>':
        # Remove rows where the column value is less than or equal to the criteria value
        df = df[df[column] <= float(value)]
    elif operation == '< =':
        # Remove rows where the column value is greater than the criteria value
        df = df[df[column] > float(value)]
    elif operation == '<':
        # The '<' case is correctly implemented in the original prompt as removing rows where the value is greater or equal
        df = df[df[column] >= float(value)]  # Correcting based on the initial explanation
    elif operation == 'contain':
        # Remove rows that contain the criteria value
        df = df[~df[column].str.contains(value, na=False, regex=True)]

    return df


def normalize_text(text):
    """ Normalize text to retain dots and commas, converting to lower case for comparison. """
    # Retaining dots, commas, and converting to lower case
    return re.sub(r'[^a-zA-Z0-9.,]+', '', text).lower()

def run_removal_automation(file_config_df, cal_df, base_path):
    base_path = Path(base_path)
    excel_writers = {}  # Manage ExcelWriter objects for each output file

    for _, file_info in file_config_df.iterrows():
        if pd.isna(file_info['input_folder_path']) or pd.isna(file_info['input_file_name']):
            continue

        input_file_path = base_path / file_info['input_folder_path'] / file_info['input_file_name']
        input_file_path = input_file_path.with_suffix(file_info['input_file_type'] if not input_file_path.suffix else input_file_path.suffix)


        output_file_name = file_info['output_file_name']
            # Check if the file name has an extension, if not, append '.xlsx'
        if '.' not in output_file_name:
            output_file_name += '.xlsx'

        output_file_path = base_path / file_info['output_folder_path'] / output_file_name

        print("Output file path:", output_file_path)

        # output_file_path = base_path / file_info['output_folder_path'] / file_info['output_file_name']
        # output_file_path = output_file_path.with_suffix('.xlsx' if not output_file_path.suffix else output_file_path.suffix)

        sheet_name = file_info.get('output_sheet_name', 'Sheet1')

        if not input_file_path.is_file():
            print(f"File not found: {input_file_path}")
            continue

        df = read_input_file_type(input_file_path, file_info['input_file_type'], file_info.get('input_sheet_name'))

        criteria_rows = cal_df[cal_df['linked_input_class'] == file_info['base_input_class']]
        if criteria_rows.empty:
            print(f"No matching criteria found for key: {file_info['base_input_class']}")
            continue

        # header_values = [normalize_text(value) for value in criteria_rows.iloc[0]['remover_rows_above_header_column_list'].split(',')]
        header_values = [normalize_text(value) for value in criteria_rows.iloc[0]['remove_rows_list'].split(',')]

        # Debug output of first row values
        print("First row values for debugging:", [normalize_text(str(x)) for x in df.iloc[0].values])

        # Directly check the first row as potential header
        first_row_values = [normalize_text(str(x)) for x in df.iloc[0].values]
        if all(header_value in first_row_values for header_value in header_values):
            header_row_index = 0
            print("First row matches header criteria, using as header.")
        else:
            # If not found in the first row, check subsequent rows
            header_row_index = None
            for index, row in df.iterrows():
                row_values = [normalize_text(str(x)) for x in row.values]
                if all(header_value in row_values for header_value in header_values):
                    header_row_index = index
                    print(f"Header found at index: {index}")
                    break

        
        if header_row_index is not None:
            new_header = df.iloc[header_row_index]  # Assuming this is the correct header
            df = df.iloc[header_row_index + 1:]  # Remove the header row and anything above it
            # df.columns = [col.strip() for col in new_header]  # Set the new header
            # Strip whitespace and convert to lower case for each header
            df.columns = [col.strip().lower() for col in new_header]  
            df.reset_index(drop=True, inplace=True)
            print("Header set successfully.")
        else:
            print("Error: Suitable header row could not be found based on the provided criteria.")
    
        # for idx, criteria in criteria_rows.iterrows():
        #     print(f"Applying criteria: {criteria['applied_column']} {criteria['criteria_to_remove_row']} {criteria['criteria_value']}")
        #     criteria_dict = {
        #         'applied_column': criteria['applied_column'].strip().lower(),
        #         'criteria_to_remove_row': criteria['criteria_to_remove_row'],
        #         'criteria_value': criteria['criteria_value']
        #     }
        #     df = apply_removal_criteria(df, criteria_dict)
        for idx, criteria in criteria_rows.iterrows():
            print(f"Applying criteria: {criteria['applied_column']} {criteria['criteria_to_remove_row']} {criteria['criteria_value']}")

            # Handle potential NaN or non-string values in 'applied_column'
            if pd.isna(criteria['applied_column']):
                print("No applied column provided. Skipping this criteria.")
                continue

            criteria_dict = {
                'applied_column': criteria['applied_column'].strip().lower(),
                'criteria_to_remove_row': criteria['criteria_to_remove_row'],
                'criteria_value': criteria['criteria_value']
            }
            df = apply_removal_criteria(df, criteria_dict)

        output_file_path.parent.mkdir(parents=True, exist_ok=True)

        # Manage ExcelWriter for output files
        if output_file_path not in excel_writers:
            if output_file_path.exists():
                excel_writers[output_file_path] = pd.ExcelWriter(output_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
            else:
                excel_writers[output_file_path] = pd.ExcelWriter(output_file_path, engine='openpyxl', mode='w')
        
        print("Excel writer status:", output_file_path in excel_writers)

        # Save the data to the specified sheet
        df.to_excel(excel_writers[output_file_path], sheet_name=sheet_name, index=False)
        print(f"Data written to {sheet_name} in {output_file_path}")

    # Close all Excel writers
    for path, writer in excel_writers.items():
        writer.close()
        print(f"Closed writer for {path}")