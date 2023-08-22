import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
from pathlib import Path
import openpyxl
from difflib import SequenceMatcher

path_to_output = Path().absolute() / 'data' / 'output' / 'pre_agp0' 
path_to_template = Path().absolute() / 'data' / 'input' / 'pre_agp0' / "Architecture Intake Review Engine Report Draft.docx"
path_to_rubric = Path().absolute() / 'data' / 'input' / 'pre_agp0' / "IIT-EA-Decision-Matrix.xlsx"

def main() -> None:
    generate_pre_agp_0_report("C:\AIRE\data\input\pre_agp0\sample_assessment.xlsx")

def generate_pre_agp_0_report(assessment_file_directory) -> None:
    
    attribute_levels, risk_score, rationales, corporate_or_cluster, initiative_name = get_assessment_data(assessment_file_directory)

    fill_in_template(initiative_name, attribute_levels, risk_score, rationales, corporate_or_cluster)

    print("-------------------------------------- Complete --------------------------------------")
        
def get_assessment_data(assessment_file_directory) -> list:
    """
    Get attribute_levels, risk_score, rationales, and corporate_or_cluster from the assessment file
    """
    workbook = openpyxl.load_workbook(assessment_file_directory, data_only=True)

    sheet = workbook['Matrix']
    
    attribute_levels = [sheet['D10'].value, sheet['D11'].value, sheet['D12'].value, sheet['D13'].value, sheet['D14'].value]
    risk_score = sheet['D15'].value
    rationales = [sheet['C10'].value, sheet['C11'].value, sheet['C12'].value, sheet['C13'].value, sheet['C14'].value]
    corporate_or_cluster = sheet['D22'].value
    initiative_name = sheet['C3'].value

    assessment_data = [attribute_levels, risk_score, rationales, corporate_or_cluster, initiative_name]

    return assessment_data

# this function should take in the array of rubric descriptions and user input arrays
# based on the arrays, an appropriate set of results and conclusion should be reached
# report should open automatically (probably remind the users to save the generated report)
def fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None:

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

    save_output(doc)

    open_output_file()

def generate_template() -> DocxTemplate:
    return DocxTemplate(path_to_template)

def get_similarity(rubric_text: str, rationale_text: str) -> float:
    """
    Calcualte the text similarity ratio between the rubric text and input text
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

    rubric_descriptions = []
    dataframe = pd.read_excel(path_to_rubric, sheet_name="Rubric")
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

def save_output(doc: DocxTemplate) -> None:
    doc.save(path_to_output / "Assessment_Report.docx")

def open_output_file() -> None:
    os.startfile(str(path_to_output))

if __name__ == "__main__":
    main()