from early_engagement.early_engagement import EarlyEngagement
from pathlib import Path

# For testing purpose
#
# 1. early_engagement module receives a path to operational plan file and generates comparison report (compare with the previous runs' operational plan file) and intake forms (to be developed)
# 2. Output will be saved in the data/archive folder 
#   Output 1: a zip file of all the outputs with the name of current datetime
#   Output 2: submitted operational plan file with the name of 'previous_op.xlsx' to be compared with the next run's operational plan
#   * all the files in the data/output folder will be deleted for the next run
# 3. In case of the first run or if the submitted operational plan is same as the previous run's one, comparison reports will not be generated
# 4. If the submitted operational plan is different form the previous one, the program will generate comparison reports only
#   Comparison report 1: Word report that has [data] dictionary variable recorded
#   Comparison report 2: Operational plan which changed cells are highlight with blue color and have previous and current values
# 5. path_to_op_2 has some changes and is different from path_to_op_2 (for testing purpose) 
# 6. to make the program generate comparison reports, change the arguemnt of early_engagement([argument]) from path_to_op_1 to path_to_op_2 or from path_to_op_1 to path_to_op_2

path_to_op_1 = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022.xlsx'
path_to_op_2 = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022 - old.xlsx'

try:
    EarlyEngagement(path_to_op_1).check_first_run().add_previous_op_to_output_folder().compare_current_previous_op().generate_comparison_report().generate_comparison_tables().generate_intake_forms().archive_files().clear_output_folder()
except Exception as e:
    print(e)