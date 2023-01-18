from pathlib import Path
import shutil
from datetime import datetime
import os
import pandas as pd
from pprint import pprint 

class early_engagement():
    '''
    __init__
    add_prev_op_path
    add_prev_curr_op_df
    add_prev_curr_op_comparison
    print_data
    archive_curr_op
    '''
    def __init__(self, path_to_curr_op = None):
        self.data = {}
        self.data['path_to_curr_op'] = path_to_curr_op

    def print_data(self):
        pprint(self.data)
        return self

    def add_prev_op_path(self):
        path_to_prev_op = self.get_prev_op_path()
        self.data['path_to_prev_op'] = path_to_prev_op
        if(path_to_prev_op == None):
            self.data['isPrev'] = False
        else:
            self.data['isPrev'] = True
        return self

    # helper function for add_prev_op
    def get_prev_op_path(self):
        path_to_ea_archive = Path().absolute() / 'data' / 'output' / 'early_engagement' / 'archive'
        file_names = []
        for path, subdirs, files in os.walk(path_to_ea_archive):
            for name in files:
                file_name = Path(os.path.join(path, name))
                file_names.append(file_name)
        if(len(file_names) == 0):
            return None
        file_names.sort()
        path_to_prev_op = file_names[-1]
        return path_to_prev_op

    def archive_curr_op(self):
        if(self.data['path_to_curr_op'] == None):
            raise Exception('archive_curr_op: No path_to_curr_op in the data')
        path_to_curr_op = self.data['path_to_curr_op']
        path_to_archive = Path().absolute() / 'data' / 'output' / 'early_engagement' / 'archive'
        current_date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        curr_op_archive_name = current_date_time + '.xlsx'
        curr_op_archive_name = curr_op_archive_name.replace('-', '').replace(':','').replace(' ','_')     # replace - and : due to file name format error
        shutil.copyfile(path_to_curr_op, path_to_archive / curr_op_archive_name)
        return self

    def add_prev_curr_op_df(self):
        if(self.data['path_to_curr_op'] == None):
            raise Exception('add_prev_curr_op_df: No path_to_curr_op in the data')
        if(self.data['path_to_prev_op'] == None):
            raise Exception('add_prev_curr_op_df: No path_to_prev_op in the data')

        self.data['prev_op_run_df'] = pd.read_excel(self.data['path_to_prev_op'], sheet_name='RUN')
        self.data['prev_op_grow_df'] = pd.read_excel(self.data['path_to_prev_op'], sheet_name='GROW')
        self.data['prev_op_transform_df'] = pd.read_excel(self.data['path_to_prev_op'], sheet_name='TRANSFORM')

        self.data['curr_op_run_df'] = pd.read_excel(self.data['path_to_curr_op'], sheet_name='RUN')
        self.data['curr_op_grow_df'] = pd.read_excel(self.data['path_to_curr_op'], sheet_name='GROW')
        self.data['curr_op_transform_df'] = pd.read_excel(self.data['path_to_curr_op'], sheet_name='TRANSFORM')

        return self

    def add_prev_curr_op_comparison(self):
        if(self.data['isPrev'] == None):
            raise Exception("add_prev_curr_op_comparison: There is no isPrev in the data.")
        if(self.data['isPrev'] == False):
            self.data['comparison'] = None
            return self.data
            
        if(self.data['prev_op_run_df'].equals(None)):
            raise Exception("add_prev_curr_op_comparison: There is no prev_op_run_df in the data.")
        if(self.data['prev_op_grow_df'].equals(None)):
            raise Exception("add_prev_curr_op_comparison: There is no prev_op_grow_df in the data.")
        if(self.data['prev_op_transform_df'].equals(None)):
            raise Exception("add_prev_curr_op_comparison: There is no prev_op_transform_df in the data.")
        if(self.data['curr_op_run_df'].equals(None)):
            raise Exception("add_prev_curr_op_comparison: There is no curr_op_run_df in the data.")
        if(self.data['curr_op_grow_df'].equals(None)):
            raise Exception("add_prev_curr_op_comparison: There is no curr_op_grow_df in the data.")
        if(self.data['curr_op_transform_df'].equals(None)):
            raise Exception("add_prev_curr_op_comparison: There is no curr_op_transform_df in the data.")

        prev_op_run_df = self.data['prev_op_run_df']
        prev_op_grow_df = self.data['prev_op_grow_df']
        prev_op_transform_df = self.data['prev_op_transform_df']
        curr_op_run_df = self.data['curr_op_run_df']
        curr_op_grow_df = self.data['curr_op_grow_df']
        curr_op_transform_df = self.data['curr_op_transform_df']

        comparison = {}

        comparison['areRunSame'] = prev_op_run_df.equals(curr_op_run_df)
        comparison['areGrowSame'] = prev_op_grow_df.equals(curr_op_grow_df)
        comparison['areTransformSame'] = prev_op_transform_df.equals(curr_op_transform_df)

        if(comparison['areRunSame'] == False):
            run_change_df = prev_op_run_df.eq(curr_op_run_df)
            comparison['run_change_df'] = run_change_df
            change_locs = run_change_df.eq(False)
            change_indexes = change_locs[change_locs].stack().index.tolist()
            comparison['run_change_indexes'] = change_indexes

        if(comparison['areGrowSame'] == False):
            grow_change_df = prev_op_grow_df.eq(curr_op_grow_df)
            comparison['grow_change_df'] = grow_change_df 
            change_locs = grow_change_df.eq(False)
            change_indexes = change_locs[change_locs].stack().index.tolist()
            comparison['grow_change_indexes'] = change_indexes

        if(comparison['areTransformSame'] == False):
            transform_change_df = prev_op_transform_df.eq(curr_op_transform_df)
            comparison['transform_change_df'] = transform_change_df
            change_locs = transform_change_df.eq(False)
            change_indexes = change_locs[change_locs].stack().index.tolist()
            comparison['transform_change_indexes'] = change_indexes

        self.data['comparison'] = comparison

        return self

    def save_excel(self):
        path_to_changes = Path().absolute() / 'data' / 'output' / 'early_engagement' / 'changes'

        if(self.data['comparison']['runChanges'] != None):
            self.data['comparison']['runChanges'].to_excel(path_to_changes / "runChanges.xlsx")
        if(self.data['comparison']['growChanges'] != None):
            self.data['comparison']['growChanges'].to_excel(path_to_changes / "growChanges.xlsx")
        if(self.data['comparison']['transformChanges'] != None):
            self.data['comparison']['transformChanges'].to_excel(path_to_changes / "transformChanges.xlsx")

        return self

# if __name__ == '__main__':
#     # For testing purpose
#     path_to_curr_op = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022.xlsx'
#     try:
#         early_engagement(path_to_curr_op).add_prev_op_path().add_prev_curr_op_df().add_prev_curr_op_comparison().print_data().save_excel()
#     except Exception as e:
#         print(str(e))
