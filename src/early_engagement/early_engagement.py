from pathlib import Path
import shutil
from datetime import datetime
import os
import pandas as pd
from pprint import pprint 
from docx import Document

class early_engagement():

    def __init__(self, path_to_current_op = None):
        self.data = {}
        self.data['path_to_current_op'] = path_to_current_op

    def add_previous_op_path(self):
        path_to_previous_op = self.get_previous_op_path()
        self.data['path_to_previous_op'] = path_to_previous_op
        
        if(path_to_previous_op == None):
            self.data['archive_exists'] = False
        else:
            self.data['archive_exists'] = True

        return self

    def archive_current_op(self):
        self.validate_path_to_current_op()

        path_to_current_op = self.data['path_to_current_op']
        path_to_archive = Path().absolute() / 'data' / 'output' / 'early_engagement' / 'archive'
        current_date_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        curr_op_archive_name = current_date_time + '.xlsx'
        curr_op_archive_name = curr_op_archive_name.replace('-', '').replace(':','').replace(' ','_')     # replace - and : due to file name format error
        shutil.copyfile(path_to_current_op, path_to_archive / curr_op_archive_name)

        return self

    def add_dataframes(self):
        self.validate_path_to_current_op()
        self.validate_path_to_previous_op()            

        self.add_dataframes_from_prev_op_run_grow_transform()
        self.add_dataframes_from_curr_op_run_grow_transform()

        return self

    def add_comparison(self):
        self.validate_archive_exists()

        if(self.data['archive_exists'] == False):
            self.data['comparison'] = None
            return self.data
            
        self.validate_previous_op_dataframes() 
        self.validate_current_op_dataframes()

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
            # add_run_changed_indexes()
            run_change_df = prev_op_run_df.eq(curr_op_run_df)
            comparison['run_change_df'] = run_change_df
            change_locs = run_change_df.eq(False)
            change_indexes = change_locs[change_locs].stack().index.tolist()
            comparison['run_change_indexes'] = change_indexes

        if(comparison['areGrowSame'] == False):
            # add_grow_change_indexes()
            grow_change_df = prev_op_grow_df.eq(curr_op_grow_df)
            comparison['grow_change_df'] = grow_change_df 
            change_locs = grow_change_df.eq(False)
            change_indexes = change_locs[change_locs].stack().index.tolist()
            comparison['grow_change_indexes'] = change_indexes

        if(comparison['areTransformSame'] == False):
            # add_transform_change_indexes()
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

    def print_data(self):
        pprint(self.data)
        return self

    def save_data_to_file(self):
        doc = Document()
        path_to_data = Path().absolute() / 'data' / 'output' / 'early_engagement' / 'data.docx'
        for key, value in self.data.items():
            doc.add_paragraph(f'{key}: {value}')
            doc.add_paragraph("-------------------------------------------------------")
        doc.save(path_to_data)
        return self

    # TBD
    # def generate_comparison_report():
    # def generate_excel_comparison_report():
    # def generate_word_comparison_report():
    # def generate_intake_forms():
    # def generate_intake_form():

    # ----------------- helper methods -----------------------

    def get_previous_op_path(self):
        path_to_archive = Path().absolute() / 'data' / 'output' / 'early_engagement' / 'archive'
        file_names = []
        for path, subdirs, files in os.walk(path_to_archive):
            for name in files:
                file_name = Path(os.path.join(path, name))
                file_names.append(file_name)
        if(len(file_names) == 0):
            return None
        file_names.sort()
        path_to_previous_op = file_names[-1]

        return path_to_previous_op

    def add_dataframes_from_prev_op_run_grow_transform(self):
        self.data['prev_op_run_df'] = pd.read_excel(self.data['path_to_previous_op'], sheet_name='RUN')
        self.data['prev_op_grow_df'] = pd.read_excel(self.data['path_to_previous_op'], sheet_name='GROW')
        self.data['prev_op_transform_df'] = pd.read_excel(self.data['path_to_previous_op'], sheet_name='TRANSFORM')

        return self;

    def add_dataframes_from_curr_op_run_grow_transform(self):
        self.data['curr_op_run_df'] = pd.read_excel(self.data['path_to_current_op'], sheet_name='RUN')
        self.data['curr_op_grow_df'] = pd.read_excel(self.data['path_to_current_op'], sheet_name='GROW')
        self.data['curr_op_transform_df'] = pd.read_excel(self.data['path_to_current_op'], sheet_name='TRANSFORM')

        return self;

    def validate_archive_exists(self):
        if(self.data['archive_exists'] == None):
            raise Exception("add_comparison: There is no archive_exists in the data.")

        return self;

    def validate_previous_op_dataframes(self):
        if(self.data['prev_op_run_df'].equals(None)):
            raise Exception("!Error - add_comparison: There is no prev_op_run_df in the data.")
        if(self.data['prev_op_grow_df'].equals(None)):
            raise Exception("!Error - add_comparison: There is no prev_op_grow_df in the data.")
        if(self.data['prev_op_transform_df'].equals(None)):
            raise Exception("!Error - add_comparison: There is no prev_op_transform_df in the data.")
  
        return self;

    def validate_current_op_dataframes(self):
        if(self.data['curr_op_run_df'].equals(None)):
            raise Exception("!Error - add_comparison: There is no curr_op_run_df in the data.")
        if(self.data['curr_op_grow_df'].equals(None)):
            raise Exception("!Error - add_comparison: There is no curr_op_grow_df in the data.")
        if(self.data['curr_op_transform_df'].equals(None)):
            raise Exception("!Error - add_comparison: There is no curr_op_transform_df in the data.")

        return self

    def validate_path_to_current_op(self):
        if(self.data['path_to_current_op'] == None):
            raise Exception('!Error - add_dataframes: No path_to_current_op in the data')

        return self;

    def validate_path_to_previous_op(self):
        if(self.data['path_to_previous_op'] == None):
            raise Exception('!Error - add_dataframes: No path_to_previous_op in the data')

        return self;

    def validate_path_to_current_op(self):
        if(self.data['path_to_current_op'] == None):
            raise Exception('!Error - archive_current_op: No path_to_current_op in the data')

        return self

if __name__ == '__main__':
    # For testing purpose
    path_to_current_op = Path().absolute() / 'data' / 'input' / 'early_engagement' / 'CYSSC FY 2022-23 Operational Plan - PUBLISHED June 2022.xlsx'
    try:
        early_engagement(path_to_current_op).add_previous_op_path().add_dataframes().add_comparison().save_data_to_file()
    except Exception as e:
        print(str(e))