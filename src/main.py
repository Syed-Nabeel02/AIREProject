from early_engagement.early_engagement import early_engagement 
from pathlib import Path

# For testing purpose
# change the argument of early_engagement(argument) - path_to_current_op or path_to_changed_op - to see different output
path_to_op_1 = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022.xlsx'
path_to_op_2 = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022 - old.xlsx'

try:
    early_engagement(path_to_op_1).check_first_run().add_previous_op_to_output_folder().compare_current_previous_op().generate_comparison_report().generate_comparison_tables().generate_intake_forms().archive_files().clear_output_folder()
except Exception as e:
    print(e)