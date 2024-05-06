import pandas as pd
import os
from pathlib import Path

def get_active_functions(base_user_path, config_file, sheet_name):
    # Load the Excel sheet into a pandas DataFrame
    file_path = base_user_path/config_file
   
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Filter the rows where 'Active' equals 1
    active_functions = df[df['Active'] == 1]['Function'].tolist()

    # Print the list of active functions
    return active_functions

# def read_excel_data_from_row(filepath, sheet_name):

#     try:
#         df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=6)

#         if "Num" in df.columns:
#             num_col_position = df.columns.get_loc("Num")
            
#             for col in df.columns:
#                 col_position = df.columns.get_loc(col)
                
#                 # If the column is to the left of or is the "Num" column, convert to string
#                 if col_position <= num_col_position:
#                     df[col] = df[col].astype(str)
#                 # If the column is to the right of "Num", apply transformations to string values
#                 else:
#                     if df[col].dtype == 'object' or df[col].dtype == 'string':
#                         df[col] = df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)
#         else:
#             print('"Num" column not found. No specific conversions based on the "Num" criteria were performed.')

#         return df
#     except FileNotFoundError:
#         print(f"File not found: {filepath}")
#         return None
#     except Exception as e:
#         print(f"An error occurred while reading {sheet_name} from {filepath}: {e}")
#         return None
    
def read_config(base_path, config_file_path, sheet_name, lower_case_except_file_zone):
    try:
        # Reading the Excel file starting from the fifth row (skip first four rows)
        config_df = pd.read_excel(config_file_path, sheet_name=sheet_name, skiprows=6)

        # Initialize empty DataFrames for the cases when columns may not exist
        file_config_df = pd.DataFrame()
        cal_df = pd.DataFrame()
        group_df = pd.DataFrame()

        # Locate necessary columns
        if "linked_input_class" in config_df.columns:
            calc_class_pos = config_df.columns.get_loc("linked_input_class")
            file_config_df = config_df.iloc[:, :calc_class_pos].copy()

            if "base_mapping_group" in config_df.columns:
                calc_pref_group_pos = config_df.columns.get_loc("base_mapping_group")
                cal_df = config_df.iloc[:, calc_class_pos:calc_pref_group_pos].copy()
                group_df = config_df.iloc[:, calc_pref_group_pos:].copy()
            else:
                # Case when only linked_input_class exists
                cal_df = config_df.iloc[:, calc_class_pos:].copy()
        else:
            # If neither linked_input_class nor base_mapping_group exist
            file_config_df = config_df.copy()

        # Transformation function
        def transform(x):
            try:
                float(x)  # Try converting to float
                return x  # Return original if it's a number
            except ValueError:
                # Apply transformation if conversion fails (meaning it's truly a string)
                if isinstance(x, str):
                    return x.lower().strip() if lower_case_except_file_zone else x.strip()
                else:
                    return x

        # Drop rows where all cells are NaN in df
        file_config_df = file_config_df.dropna(how='all')
        cal_df = cal_df.dropna(how='all')
        group_df = group_df.dropna(how='all')


        # Applying transformations where necessary
        file_config_df = file_config_df.applymap(str)
        cal_df = cal_df.applymap(transform)
        group_df = group_df.applymap(transform)

        # Adding full paths if columns exist
        for df in [file_config_df, cal_df, group_df]:
            if 'input_folder_path' in df.columns:
                df['full_input_folder_path'] = df['input_folder_path'].apply(lambda x: os.path.join(base_path, x))
            if 'output_folder_path' in df.columns:
                df['full_output_folder_path'] = df['output_folder_path'].apply(lambda x: os.path.join(base_path, x))

        return file_config_df, cal_df, group_df
    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    

    
def get_cal(file_info, cal_df):   
    filtered_df = cal_df[cal_df['linked_input_class'] == file_info['base_input_class']]
    return filtered_df


def get_group(cal_info, group_df):  
    filtered_df = group_df[group_df['base_mapping_group'] == cal_info['linked_mapping_group']]
    return filtered_df




# def read_config_except_columns(base_path, config_file_path, sheet_name, lower_case_except_file_zone, 
#                  execpt_cal_df_columns=None, execpt_group_df_columns=None):
#     try:
#         config_df = pd.read_excel(config_file_path, sheet_name=sheet_name, skiprows=5)

#         # Initialize DataFrames
#         file_config_df, cal_df, group_df = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

#         # Determine columns
#         if "linked_input_class" in config_df.columns:
#             calc_class_pos = config_df.columns.get_loc("linked_input_class")
#             file_config_df = config_df.iloc[:, :calc_class_pos].copy()

#             if "base_mapping_group" in config_df.columns:
#                 calc_pref_group_pos = config_df.columns.get_loc("base_mapping_group")
#                 cal_df = config_df.iloc[:, calc_class_pos:calc_pref_group_pos].copy()
#                 group_df = config_df.iloc[:, calc_pref_group_pos:].copy()
#             else:
#                 cal_df = config_df.iloc[:, calc_class_pos:].copy()
#         else:
#             file_config_df = config_df.copy()

#         # Apply transformations
#         file_config_df = file_config_df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
#         for col in ['base_input_class']:  # Add more columns as needed
#             file_config_df[col] = file_config_df[col].str.lower().str.strip()

#         # Apply conditional transformation for cal_df based on the parameter
#         for col in cal_df.columns:
#             if execpt_cal_df_columns and col in execpt_cal_df_columns:
#                 cal_df[col] = cal_df[col].apply(lambda x: x if isinstance(x, str) else x)
#             else:
#                 cal_df[col] = cal_df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

#         # Apply conditional transformation for group_df based on the parameter
#         for col in group_df.columns:
#             if execpt_group_df_columns and col in execpt_group_df_columns:
#                 group_df[col] = group_df[col].apply(lambda x: x if isinstance(x, str) else x)
#             else:
#                 group_df[col] = group_df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

#         # Drop NaN rows
#         file_config_df.dropna(how='all', inplace=True)
#         cal_df.dropna(how='all', inplace=True)
#         group_df.dropna(how='all', inplace=True)

#         return file_config_df, cal_df, group_df
#     except Exception as e:
#         print(f"An error occurred: {e}")
#         return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# def read_config_except_columns(base_path, config_file_path, sheet_name, lower_case_except_file_zone, 
#                  execpt_cal_df_columns=None, execpt_group_df_columns=None):
#     try:
#         config_df = pd.read_excel(config_file_path, sheet_name=sheet_name, skiprows=6)

#         # Initialize DataFrames
#         file_config_df, cal_df, group_df = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

#         # Determine columns
#         if "linked_input_class" in config_df.columns:
#             calc_class_pos = config_df.columns.get_loc("linked_input_class")
#             file_config_df = config_df.iloc[:, :calc_class_pos].copy()

#             if "base_mapping_group" in config_df.columns:
#                 calc_pref_group_pos = config_df.columns.get_loc("base_mapping_group")
#                 cal_df = config_df.iloc[:, calc_class_pos:calc_pref_group_pos].copy()
#                 group_df = config_df.iloc[:, calc_pref_group_pos:].copy()
#             else:
#                 cal_df = config_df.iloc[:, calc_class_pos:].copy()
#         else:
#             file_config_df = config_df.copy()

#         # Apply transformations
#         file_config_df = file_config_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
#         for col in ['base_input_class']:  # Add more columns as needed
#             file_config_df[col] = file_config_df[col].str.lower().str.strip()

#         # Apply conditional transformation for cal_df based on the parameter
#         for col in cal_df.columns:
#             if execpt_cal_df_columns and col in execpt_cal_df_columns:
#                 cal_df[col] = cal_df[col].apply(lambda x: x if isinstance(x, str) else x)
#             else:
#                 cal_df[col] = cal_df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

#         # Apply conditional transformation for group_df based on the parameter
#         for col in group_df.columns:
#             if execpt_group_df_columns and col in execpt_group_df_columns:
#                 group_df[col] = group_df[col].apply(lambda x: x if isinstance(x, str) else x)
#             else:
#                 group_df[col] = group_df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

#         # Drop NaN rows
#         file_config_df.dropna(how='all', inplace=True)
#         cal_df.dropna(how='all', inplace=True)
#         group_df.dropna(how='all', inplace=True)

#         return file_config_df, cal_df, group_df
#     except Exception as e:
#         print(f"An error occurred: {e}")
#         return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def read_config_except_columns(base_path, config_file_path, sheet_name, lower_case_except_file_zone, 
                               execpt_cal_df_columns=None, execpt_mapping_zone_df_columns=None):
    try:
        config_df = pd.read_excel(config_file_path, sheet_name=sheet_name, skiprows=6)

        # Initialize DataFrames
        file_config_df, cal_df, mapping_zone_df = pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

        # Determine columns
        if "linked_input_class" in config_df.columns:
            calc_class_pos = config_df.columns.get_loc("linked_input_class")
            file_config_df = config_df.iloc[:, :calc_class_pos].copy()

            if "base_mapping_group" in config_df.columns:
                calc_pref_group_pos = config_df.columns.get_loc("base_mapping_group")
                cal_df = config_df.iloc[:, calc_class_pos:calc_pref_group_pos].copy()
                mapping_zone_df = config_df.iloc[:, calc_pref_group_pos:].copy()
            else:
                cal_df = config_df.iloc[:, calc_class_pos:].copy()
        else:
            file_config_df = config_df.copy()

        # Only strip column headers after they exist
        if not file_config_df.empty:
            file_config_df.columns = file_config_df.columns.str.strip()
        if not cal_df.empty:
            cal_df.columns = cal_df.columns.str.strip()
        if not mapping_zone_df.empty:
            mapping_zone_df.columns = mapping_zone_df.columns.str.strip()

        # Apply transformations
        file_config_df = file_config_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
        file_config_df['base_input_class'] = file_config_df['base_input_class'].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

        # Apply conditional transformation for cal_df
        for col in cal_df.columns:
            if execpt_cal_df_columns and col in execpt_cal_df_columns:
                # Only strip spaces from string values in excepted columns
                cal_df[col] = cal_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            else:
                # Strip spaces and convert to lowercase for other columns
                cal_df[col] = cal_df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

        # Apply conditional transformation for mapping_zone_df
        for col in mapping_zone_df.columns:
            if execpt_mapping_zone_df_columns and col in execpt_mapping_zone_df_columns:
                # Only strip spaces from string values in excepted columns
                mapping_zone_df[col] = mapping_zone_df[col].apply(lambda x: x.strip() if isinstance(x, str) else x)
            else:
                # Strip spaces and convert to lowercase for other columns
                mapping_zone_df[col] = mapping_zone_df[col].apply(lambda x: x.lower().strip() if isinstance(x, str) else x)

        # Drop NaN rows
        file_config_df.dropna(how='all', inplace=True)
        cal_df.dropna(how='all', inplace=True)
        mapping_zone_df.dropna(how='all', inplace=True)

        return file_config_df, cal_df, mapping_zone_df
    except Exception as e:
        print(f"An error occurred: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()



def read_excel_data_from_row(filepath, sheet_name):
    """
    Reads data from an Excel sheet beginning at cell A3.

    Parameters:
    - filepath: str or Path, the path to the Excel file.
    - sheet_name: str, the name of the sheet to read data from.

    Returns:
    - DataFrame containing the data from the specified sheet starting at row 2.
    """
    try:
        # Read the data from the specified sheet starting at row 2
        # skiprows=3 skips the first two rows (0-indexed), starting the read from row 2
        df = pd.read_excel(filepath, sheet_name=sheet_name, skiprows=6)
        return df
    except FileNotFoundError:
        print(f"File not found: {filepath}")
        return None
    except Exception as e:
        print(f"An error occurred while reading {sheet_name} from {filepath}: {e}")
        return None     


def set_config_df_types(config_df):
    """
    Sets the correct data types for the columns in the config DataFrame.
    """
    dtype_columns = ['input_folder_path', 'input_file_name', 'input_file_type',
                     'output_folder_path', 'output_file_name']
    for col in dtype_columns:
        config_df[col] = config_df[col].astype(str)
    return config_df

def set_config_df_types_2(config_df):
    """
    Sets the correct data types for the columns in the config DataFrame.
    """
    dtype_columns = ['input_folder_path', 'input_file_name', 'input_file_type',
                     'output_folder_path', 'output_file_name_txt', 'output_file_name_xlsx']
    for col in dtype_columns:
        config_df[col] = config_df[col].astype(str)
    return config_df

def set_config_df_types_3(config_df):
    """
    Sets the correct data types for the columns in the config DataFrame.
    """
    dtype_columns = ['input_folder_path', 'input_file_name', 'input_file_type',
                     'output_folder_path', 'output_file_name']
    for col in dtype_columns:
        config_df[col] = config_df[col].astype(str)
    return config_df

def get_active_functions_sorted(base_user_path, config_file, sheet_name):
    # Construct the file path
    file_path = base_user_path / config_file
   
    # Load the Excel sheet into a pandas DataFrame
    df = pd.read_excel(file_path, sheet_name=sheet_name)

    # Filter the rows where 'Active' equals 1
    active_functions = df[df['Active'] == 1]
    
    # Sort the filtered DataFrame by the 'Order' column in ascending order
    sorted_active_functions = active_functions.sort_values(by='Order')
    
    # Convert the sorted DataFrame into a list of dictionaries
    active_functions_list = sorted_active_functions[['Function', 'Active', 'Order', 'Time.sleep (s)']].to_dict('records')

    # Return the sorted list of active functions
    return active_functions_list