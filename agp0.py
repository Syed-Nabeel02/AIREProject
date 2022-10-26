import os
import shutil
from pathlib import Path
import zipfile
from tkinter import messagebox

"""
AGP0: Submit a zip file of supplement files and check whether all requirements are met.

Methods
- generate_agp_0_report() -> None
- extract_zip_file(zip_file_name: str) -> None
- get_file_names_in_folder() -> list
- organize_files_by_type(file_names) -> tuple
"""

def generate_agp_0_report(agp0_supplements_directory) -> None:

    extract_zip_file(agp0_supplements_directory)

    file_names = get_file_names_in_folder()

    word_file_names, excel_file_names, ppt_file_names, pdf_file_names = organize_files_by_type(file_names)

    requirements = ["SAS", "PAQ", "AR Log", "Decision Matrix"]
    checked = []

    for word_file_name in word_file_names:
        checked.append(check_requirements(word_file_name))

    for excel_file_name in excel_file_names:
        checked.append(check_requirements(excel_file_name))

    for ppt_file_name in ppt_file_names:
        checked.append(check_requirements(ppt_file_name))

    for pdf_file_name in pdf_file_names:
        checked.append(check_requirements(pdf_file_name))

    print(checked)
    
    result = "SAS, PAQ, AR Log, and Decision Matrix are required.\n"
    for check in checked:
        for c in check:
            result = result + c

    messagebox.askokcancel("Result", result)

    base_dir = Path(__file__).parent
    directory_to_extract_to = base_dir / "AGP0 Supplements"
    shutil.rmtree(directory_to_extract_to)

def check_requirements(file_name: str) -> list:
    file_name_original = file_name
    file_name = file_name.lower()

    checked = []
    
    if 'sas' in file_name: # the project name can be 'sash'. needs improvement. naming conventions.
        checked.append("SAS exists!: " + file_name_original)
    if 'paq' in file_name:
        checked.append("PAQ exists!: " + file_name_original)
    if 'ar' in file_name and 'log' in file_name:
        checked.append("AR Log exists!: " + file_name_original)
    if 'decision' in file_name and 'matrix' in file_name:
        checked.append("Decision Matrix exists!: " + file_name_original)
    
    return checked

def extract_zip_file(agp0_supplements_directory: str) -> None:
    base_dir = Path(__file__).parent
    path_to_zip_file = agp0_supplements_directory
    directory_to_extract_to = base_dir / "AGP0 Supplements"

    with zipfile.ZipFile(path_to_zip_file, 'r') as zip_ref:
        zip_ref.extractall(directory_to_extract_to)

def get_file_names_in_folder() -> list:
    file_names = []
    base_dir = Path(__file__).parent
    directory_to_extract_to = base_dir / "AGP0 Supplements"
    for path, subdirs, files in os.walk(directory_to_extract_to):
        for name in files:
            file_name = Path(os.path.join(path, name)).name
            file_names.append(file_name)
    return file_names

def organize_files_by_type(file_names) -> tuple:
    word_file_names = []
    excel_file_names = []
    ppt_file_names = []
    pdf_file_names = []

    for file_name in file_names:
        if file_name[-4:] == "docx":
            word_file_names.append(file_name)
        elif file_name[-4:] == "xlsx":
            excel_file_names.append(file_name)
        elif file_name[-4:] == "pptx":
            ppt_file_names.append(file_name)
        elif file_name[-3:] == "pdf":
            pdf_file_names.append(file_name)
    
    return (word_file_names, excel_file_names, ppt_file_names, pdf_file_names)