from early_engagement.early_engagement import early_engagement 
from pathlib import Path

# path to a testing purpose operational file
path_to_current_op = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022.xlsx'

early_engagement(path_to_current_op).add_previous_op_path().add_dataframes().add_comparison().print_data()