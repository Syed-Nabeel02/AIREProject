from datetime import datetime
import pandas as pd
import os.path
from docxtpl import DocxTemplate
from pathlib import Path
import numpy as np
from alive_progress import alive_bar
from pprint import pprint
import json

# line 82-89 : what to do with old files **********************************

"""
Creates an output folder and many intake form Word files with proper names based on an operational plan Excel file. 

1.Template Word file (fixed)
2.Operation plan Excel file (prompt)
->
1.'OUTPUT' folder - Several filled template files for each item in an Operational plan Excel file worksheet (RUN or GROW or TRANSFORM)
2.'NEWER OUTPUT' folder - empty

Methods
----------
run(operational_plan_path: str, sheet_name: str) -> None
generate_templates(directory: str, sheet_name: str) -> None
fill_in_template(doc: DocxTemplate, record:dict, date_time: datetime) -> DocxTemplate
generate_file_name(record: dict, sheet_name: str) -> str
open_output_folder() -> None
"""

def run(operational_plan_path: str, sheet_name: str) -> None:
    """
    Does the magic

    Parameters
    ----------
    directory: str
        a directory to an operational plan file that the intake forms will be based on
    sheet_name: str
        one of the type of sheets - Run, GROW, or TRANSFORM
    """
    generate_templates(operational_plan_path, sheet_name=sheet_name)

def generate_templates(operational_plan_path: str, sheet_name: str) -> None:
    """
    Generate Word template files from an Excel file.
    
    Parameters
    ----------
    sheet_name : str
    """

    intake_form_template_path = Path(__file__).parent.parent.parent / "resources" / "early_engagement" / "EA Engagement Self-Assessment Template v0.6.docx"

    # CREATE OUTPUT FOLDER FOR WORD DOCS
    # exists_ok is for just in case a folder with that name exists already
    early_engagement_output_folder_path = Path(__file__).parent.parent.parent / "output" / "early_engagement_output"
    early_engagement_output_folder_path.mkdir(exist_ok=True)

    today = datetime.now()
    date_time = today.strftime("%m/%d/%Y, %H:%M:%S")

    dataframe = pd.read_excel(operational_plan_path, sheet_name) # many items as a dataframe

    print("-------------------------- Start ---------------------------------")

    for index, record in enumerate(dataframe.to_dict(orient="records")[:]): # for each operational plan item

        # 1.Generate and fill in a template for an item
        doc = DocxTemplate(intake_form_template_path)
        doc = fill_in_template(doc, record, date_time)

        # 2.Generate a name for the intake form
        name = generate_file_name(record, sheet_name)

        # ************************************************
        # 3.Save the intake form at the correct directory
        early_engagement_output_file_path = early_engagement_output_folder_path / name
        if not os.path.isfile(early_engagement_output_file_path):
            # If file does not exist, save to output path, this line allows for copies to not be made when ran
            name = "NEW_" + name
            early_engagement_output_file_path = early_engagement_output_folder_path / name
            doc.save(early_engagement_output_file_path) 
        else: # if the file already exists, do not save the file
            continue

    open_output_folder()
    
    print("-------------------------- Complete ---------------------------------")

def fill_in_template(doc: DocxTemplate, record:dict, date_time: datetime) -> DocxTemplate:
    """
    Fill in the template with an item's details
    """
    doc.render(record) 
    doc.add_paragraph("This was Autogenerated on" + " " + date_time)
    return doc

def generate_file_name(record: dict, sheet_name: str) -> str:
    """
    Generate a file name based on an item's details and sheet_name
    """
    IDColumn = record['ID']
    InitColumn = record['Initiative']
    ItemCol = record['WorkItemName']
    BranchColumn = record['AccountableBranch'][-6:]

    BranchColumn = BranchColumn.strip()
    InitColumn = InitColumn.strip()
    ItemCol = ItemCol.strip()

    if record['MustDoCantFail'] == 'Yes':
        name = f"{IDColumn}{'_'}{BranchColumn}{'_'}{'MDCF'}{'_'}{sheet_name[0]}{'_'}{InitColumn}{'_'}{ItemCol}"
    else:
        name = f"{IDColumn}{'_'}{BranchColumn}{'_'}{sheet_name[0]}{'_'}{InitColumn}{'_'}{ItemCol}"
    
    name = name.replace(' ', '')
    name = name.replace('-', '')
    name = name.replace('.', '')
    name = name + ".docx"
    
    return name

def open_output_folder() -> None:
    early_engagement_output_folder_path = Path(__file__).parent.parent.parent / "output" / "early_engagement_output"
    os.system("start " + str(early_engagement_output_folder_path))