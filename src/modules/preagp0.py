import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
from pathlib import Path
import openpyxl
from difflib import SequenceMatcher

"""
Pre-AGP0: Creates a report Word file based on a rubric Excel file. 

Methods
----------
run(initiative_name, assessment_file_directory) -> None

generate_pre_agp_0_report(initiative_name, assessment_file_directory) -> None

get_assessment_data(assessment_file_directory) -> list
split_assessment_data(assessment_data) -> tuple

get_similarity(rubric_text: str, rationale_text: str) -> float
get_rubric_descriptions() -> list
get_today_date_time() -> datetime
get_risk_level(risk_score: int) -> str
get_attribute_score(attribute_level: str) -> int
generate_output_folder() -> None
generate_template(output_template_name: str) -> DocxTemplate
save_output_template(doc: DocxTemplate) -> None
open_output_file() -> None
fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None
"""

def run(initiative_name: str, assessment_file_directory: str) -> None:
    """
    Does the magic

    Parameters
    ----------
    initiative_name: str
        the initiative name that is going to be written in the report
    assessment_file_directory: str
        a directory of an assessment file that the report is going to be based on
    """
    print("pre_agp0 module running")
    generate_pre_agp_0_report(initiative_name=initiative_name, assessment_file_directory=assessment_file_directory)

def generate_pre_agp_0_report(initiative_name, assessment_file_directory) -> None:

    assessment_data = get_assessment_data(assessment_file_directory)
    attribute_levels, risk_score, rationales, corporate_or_cluster = split_assessment_data(assessment_data)
    
    fill_in_template(initiative_name, attribute_levels, risk_score, rationales, corporate_or_cluster)

    print("-------------------------------------- Complete --------------------------------------")

def get_assessment_data(assessment_file_directory) -> list:
    """
    Get attribute_levels, risk_score, rationales, and corporate_or_cluster from the assessment file
    """
    excel_path = assessment_file_directory
    workbook = openpyxl.load_workbook(excel_path, data_only=True)

    sheet = workbook['Matrix']
    
    attribute_levels = [sheet['D10'].value, sheet['D11'].value, sheet['D12'].value, sheet['D13'].value, sheet['D14'].value]
    risk_score = sheet['D15'].value
    rationales = [sheet['C10'].value, sheet['C11'].value, sheet['C12'].value, sheet['C13'].value, sheet['C14'].value]
    corporate_or_cluster = sheet['D22'].value
    assessment_data = [attribute_levels, risk_score, rationales, corporate_or_cluster]

    return assessment_data

def split_assessment_data(assessment_data) -> tuple:
    attribute_levels = assessment_data[0]
    risk_score = assessment_data[1]
    rationales = assessment_data[2]
    corporate_or_cluster = assessment_data[3]
    return (attribute_levels, risk_score, rationales, corporate_or_cluster)

# this function should take in the array of rubric descriptions and user input arrays
# based on the arrays, an appropriate set of results and conclusion should be reached
# report should open automatically (probably remind the users to save the generated report)
def fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None:

    generate_output_folder()

    doc = generate_template()

    if risk_score < 7:
        gov = 'does not'
    else:
        gov = 'does'

    rubric_descriptions = get_rubric_descriptions() # ********************************

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

    open_output_file()

def generate_output_folder() -> None:
    output_folder_path = Path(__file__).parent.parent.parent / "output" / "pre_agp0_output"
    output_folder_path.mkdir(exist_ok=True)

def generate_template() -> DocxTemplate:
    template_path = Path(__file__).parent.parent.parent / "resources" / "pre_agp0" / "Architecture Intake Review Engine Report Draft.docx"
    doc = DocxTemplate(template_path)
    return doc

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
    # base_dir = Path(__file__).parent
    # # word_doc = base_dir / "Pre_AGP0" / "Architecture Intake Review Engine Report Draft.docx"
    # output_dir = base_dir / "Pre_AGP0_Output"
    # output_dir.mkdir(exist_ok=True)
    # # doc = DocxTemplate(word_doc)

    template_path = Path(__file__).parent.parent.parent / "resources" / "pre_agp0" / "IIT-EA-Decision-Matrix.xlsx"
    
    rubric_descriptions = []
    dataframe = pd.read_excel(template_path, sheet_name="Rubric")
    for record in dataframe.to_dict(orient="records"):
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

def save_output_template(doc: DocxTemplate) -> None:
    output_file_path = Path(__file__).parent.parent.parent / "output" / "pre_agp0_output" / "pre_agp0_assessment_report.docx"
    doc.save(output_file_path)

def open_output_file() -> None:
    output_file_path = Path(__file__).parent.parent.parent / "output" / "pre_agp0_output" / "pre_agp0_assessment_report.docx"
    os.system("start " + str(output_file_path))

if __name__ == '__main__':
    base_dir = Path(__file__).parent.parent.parent / "resources" / "pre_agp0"
    print(base_dir)