import pandas as pd
import os
import sys
from pathlib import Path
from f00_read_configs import read_config_except_columns

# Setup the base path and files
base_path = Path('/usr/src/app')
config_file = 'read_workflow_config_test.xlsx'
sheet_name = 'F001_(4)'
config_file_path = base_path / config_file

# Read configuration and other data
file_config_df, cal_df, mapping_zone_df = read_config_except_columns(
    base_path, config_file_path, sheet_name,
    lower_case_except_file_zone=True,
    execpt_cal_df_columns={"criteria_value"}, 
    execpt_mapping_zone_df_columns=None
)

from f01_remove_rows import read_input_file_type, apply_removal_criteria, normalize_text, run_removal_automation

# Execute the removal automation and print results
results = run_removal_automation(file_config_df, cal_df, base_path)

# Optionally print the results if the function returns something
if results:
    print(results)
