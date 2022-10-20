from tokenize import String
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
from pathlib import Path
import openpyxl
from difflib import SequenceMatcher
from pprint import pprint
import zipfile

"""
1) Pre-AGP0: Creates a report Word file based on a rubric Excel file. 
2) AGP0: Submit a zip file of supplement files and check whether all requirements are met.

Methods
- open_home_menu() -> None

- generate_pre_agp_0_report() -> None
- get_assessment_data() -> list
- get_similarity(rubric_text: str, rationale_text: str) -> float
- get_rubric_descriptions() -> list
- get_today_date_time() -> datetime
- get_risk_level(risk_score: int) -> str
- get_attribute_score(attribute_level: str) -> int
- fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None
- check_file_exist(file_name: str) -> None
- generate_output_folder() -> None:
- generate_output_template(output_template_name: str) -> DocxTemplate:
- save_output_template(doc: DocxTemplate) -> None:
- open_output_template() -> None:

- generate_agp_0_report() -> None
- extract_zip_file(zip_file_name: str) -> None
- get_file_names_in_folder() -> list
- organize_files_by_type(file_names) -> tuple
"""


def generate_pre_agp_0_report(initiative_name, assessment_file_directory) -> None:

    assessment_data = get_assessment_data(assessment_file_directory)

    attribute_levels = assessment_data[0]
    risk_score = assessment_data[1]
    rationales = assessment_data[2]
    corporate_or_cluster = assessment_data[3]

    fill_in_template(initiative_name, attribute_levels, risk_score, rationales, corporate_or_cluster)

    print("-------------------------------------- Complete --------------------------------------")

def get_assessment_data(assessment_file_directory) -> list:
    """
    Get attribute_levels, risk_score, rationales, and corporate_or_cluster from the assessment file
    """
    excel_path = assessment_file_directory

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb['Matrix']
    
    attribute_levels = [sheet['D10'].value, sheet['D11'].value, sheet['D12'].value, sheet['D13'].value, sheet['D14'].value]
    risk_score = sheet['D15'].value
    rationales = [sheet['C10'].value, sheet['C11'].value, sheet['C12'].value, sheet['C13'].value, sheet['C14'].value]
    corporate_or_cluster = sheet['D22'].value

    assessment_data = [attribute_levels, risk_score, rationales, corporate_or_cluster]

    return assessment_data

def get_similarity(rubric_text: str, rationale_text: str) -> float:
    """
    Get the text similarity ratio between the rubric text and input text
    """
    return SequenceMatcher(None, rubric_text, rationale_text).ratio()


def get_rubric_descriptions() -> list:
    """
    Get rubric description 0~12 from "IIT-EA-Decision-Matrix.xlsx" excel file
    business_scopes: 0,1,2
    it_solution: 3,4,5
    technology_upgrade: 6,7,8
    info_requirements: 9,10,11
    info_sensitivy: 12,13,14
    """
    base_dir = Path(__file__).parent
    # word_doc = base_dir / "Pre_AGP0" / "Architecture Intake Review Engine Report Draft.docx"
    output_dir = base_dir / "Pre_AGP0_Output"
    output_dir.mkdir(exist_ok=True)
    # doc = DocxTemplate(word_doc)

    excel_path = base_dir / "Pre_AGP0" / "IIT-EA-Decision-Matrix.xlsx"
    df = pd.read_excel(excel_path, sheet_name="Rubric")
    rubric_descriptions = []
    for record in df.to_dict(orient="records"):
        c = record['Description']
        rubric_descriptions.append(c)
        
    return rubric_descriptions

def get_today_date_time() -> datetime:
    today = datetime.now()
    date_time = today.strftime("%m/%d/%Y, %H:%M:%S")
    return date_time

def get_risk_level(risk_score: int) -> str:
    """
    risk_score -> risk_level: Low - 7 - Medium - 11 - High 
    """
    if risk_score < 7:
        risk_level = 'Low'
    elif 7 <= risk_score < 11:
        risk_level = 'Medium'
    else:
        risk_level = 'High'
    return risk_level

def get_attribute_score(attribute_level: str) -> int:
    """
    attribute_level -> attribute_score : Low -> 0, Medium -> 1, High -> 2
    attributes: business_scopes, it_solution, technology_upgrade, info_requirements, info_sensitivy
    """
    if attribute_level == 'Low':
        return 0
    elif attribute_level == 'Medium':
        return 1
    else:
        return 2

def check_file_exist(file_name: str) -> None:
    if os.path.exists(file_name):
        print(file_name, ' file exists')
    else:
        print(file_name, ' file does not exist')

def generate_output_folder() -> None:
    base_dir = Path(__file__).parent
    output_dir = base_dir / "Pre_AGP0_Output"
    output_dir.mkdir(exist_ok=True)

def generate_output_template(output_template_name: str) -> DocxTemplate:
    base_dir = Path(__file__).parent
    word_doc = base_dir / "Pre_AGP0" / output_template_name
    doc = DocxTemplate(word_doc)
    return doc

def save_output_template(doc: DocxTemplate) -> None:
    base_dir = Path(__file__).parent
    output_dir = base_dir / "Pre_AGP0_Output"
    output_path = output_dir / "generated_doc.docx"
    doc.save(output_path)

def open_output_template() -> None:
    base_dir = Path(__file__).parent
    output_dir = base_dir / "Pre_AGP0_Output"
    output_path = output_dir / "generated_doc.docx"
    os.system("start " + str(output_path))

# this function should take in the array of rubric descriptions and user input arrays
# based on the arrays, an appropriate set of results and conclusion should be reached
# report should open automatically (probably remind the users to save the generated report)
def fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None:

    generate_output_folder()

    doc = generate_output_template("Architecture Intake Review Engine Report Draft.docx")

    if risk_score < 7:
        gov = 'does not'
    else:
        gov = 'does'

    rubric_descriptions = get_rubric_descriptions() 
    business_scopes_rubrics = [rubric_descriptions[0], rubric_descriptions[1], rubric_descriptions[2]]
    it_solution_rubrics = [rubric_descriptions[3], rubric_descriptions[4], rubric_descriptions[5]]
    technology_upgrade_rubrics = [rubric_descriptions[6], rubric_descriptions[7], rubric_descriptions[8]]
    info_requirements_rubrics = [rubric_descriptions[9], rubric_descriptions[10], rubric_descriptions[11]]
    info_sensitivy_rubrics = [rubric_descriptions[12], rubric_descriptions[13], rubric_descriptions[14]]

    context = {'date': get_today_date_time(),
               'initiative': initiative_name,
               'score': risk_score,
               'risk': get_risk_level(risk_score),
               'gov': gov,
               'ca': attribute_levels[0],
               'ca1': attribute_levels[1],
               'ca2': attribute_levels[2],
               'ca3': attribute_levels[3],
               'ca4': attribute_levels[4],
               'rational': rationales[0],
               'rational1': rationales[1],
               'rational2': rationales[2],
               'rational3': rationales[3],
               'rational4': rationales[4],
               'comp': business_scopes_rubrics[get_attribute_score(attribute_levels[0])],
               'comp1': it_solution_rubrics[get_attribute_score(attribute_levels[1])],
               'comp2': technology_upgrade_rubrics[get_attribute_score(attribute_levels[2])],
               'comp3': info_requirements_rubrics[get_attribute_score(attribute_levels[3])],
               'comp4': info_sensitivy_rubrics[get_attribute_score(attribute_levels[4])],
               'cluster_corporate': corporate_or_cluster,
               's': str(round((get_similarity(business_scopes_rubrics[get_attribute_score(attribute_levels[0])], rationales[0]) * 100), 2)) + '%',
               's1': str(round((get_similarity(it_solution_rubrics[get_attribute_score(attribute_levels[1])], rationales[1]) * 100), 2)) + '%',
               's2': str(round((get_similarity(technology_upgrade_rubrics[get_attribute_score(attribute_levels[2])], rationales[2]) * 100), 2)) + '%',
               's3': str(round((get_similarity(info_requirements_rubrics[get_attribute_score(attribute_levels[3])], rationales[3]) * 100), 2)) + '%',
               's4': str(round((get_similarity(info_sensitivy_rubrics[get_attribute_score(attribute_levels[4])], rationales[4]) * 100), 2)) + '%',
               }

    doc.render(context)

    save_output_template(doc)

    open_output_template()

def generate_agp_0_report() -> None:
    print("1.Input a zip file of all artifact and supplement files (decision matrix, PAQ, SAS and etc)")

    zip_file_name = input("Enter the zip file name (without .zip): ")
    extract_zip_file(zip_file_name)

    print("2.Check the completeness")
    file_names = get_file_names_in_folder()
    word_file_names, excel_file_names, ppt_file_names, pdf_file_names = organize_files_by_type(file_names)

    pprint(word_file_names)
    pprint(excel_file_names)
    pprint(ppt_file_names)
    pprint(pdf_file_names)

def extract_zip_file(zip_file_name: str) -> None:
    zip_file_name = zip_file_name + ".zip"

    base_dir = Path(__file__).parent
    path_to_zip_file = base_dir / "Put an AGP0 Supplements zip file here" / zip_file_name
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

