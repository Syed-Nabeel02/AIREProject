from tokenize import String
import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
from pathlib import Path
import openpyxl
from difflib import SequenceMatcher

"""
    Fill in the report for us
"""
# this  function welcomes the user to the program
def open_home_menu():
    print("WELCOME TO OUR ARCHITECTURE INTAKE REVIEW ENGINE BOT!")
    input("Press enter to continue :)")

    invalid = True
    answer = 0

    while invalid:
        print("Please select which task you'd like to perform:\n")
        print("1. Pre AGP 0 Assessment: Use I&IT Decision Matrix to better assess a project item")
        print("2. AGP 0 Assessment: Ensure all mandatory files are submitted by Project group")
        print("3. Exit Program")
        answer = input()
        if 1 <= int(answer) <= 3:
            invalid = False

    if int(answer) == 1:
        initiative_name = input("What is the initiative name? (any name):")
        get_pre_AGP_0_report(initiative_name)
    elif int(answer) == 2:
        print("--- In Development ---")
        input("Press enter to exit")
        exit()
    else:
        exit()

    return

def get_pre_AGP_0_report(initiative_name: str) -> int:
    filename = input("Please enter the name of the assessment file (userinput.xlsx):")

    base_dir = Path(__file__).parent
    excel_path = base_dir / filename

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb['Matrix']

    attribute_levels = [sheet['D10'].value, sheet['D11'].value, sheet['D12'].value, sheet['D13'].value, sheet['D14'].value]
    risk_score = sheet['D15'].value
    rationales = [sheet['C10'].value, sheet['C11'].value, sheet['C12'].value, sheet['C13'].value, sheet['C14'].value]
    corporate_or_cluster = sheet['D22'].value

    fill_in_template(attribute_levels, risk_score, rationales, initiative_name, corporate_or_cluster)

    return 0

def get_similarity(rubric_text :str, input_text: str) -> float:
    return SequenceMatcher(None, rubric_text, input_text).ratio()


def get_rubric_descriptions() -> list:
    """
    Load data from the rubric "IIT-EA-Decision-Matrix.xlsx"
    """
    base_dir = Path(__file__).parent
    # word_doc = base_dir / "Architecture Intake Review Engine Report Draft.docx"
    output_dir = base_dir / "Output"
    output_dir.mkdir(exist_ok=True)
    # doc = DocxTemplate(word_doc)

    excel_path = base_dir / "IIT-EA-Decision-Matrix.xlsx"
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
    if risk_score < 7:
        risk_level = 'Low'
    elif 7 <= risk_score < 11:
        risk_level = 'Medium'
    else:
        risk_level = 'High'
    return risk_level


def get_attribute_score(attribute_level: str) -> int:
    if attribute_level == 'Low':
        return 0
    elif attribute_level == 'Medium':
        return 1
    else:
        return 2

# this function should take in the array of rubric descriptions and user input arrays
# based on the arrays, an appropriate set of results and conclusion should be reached
# report should open automatically (probably remind the users to save the generated report)
def fill_in_template(attribute_levels: list, risk_score: int, rationales: list, initiative_name: str, corporate_or_cluster: str) -> int:

    base_dir = Path(__file__).parent
    word_doc = base_dir / "Architecture Intake Review Engine Report Draft.docx"
    output_dir = base_dir / "Output"
    output_dir.mkdir(exist_ok=True)

    doc = DocxTemplate(word_doc)
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

    if os.path.exists('demo1.docx'):
        print('True')
    else:
        print('False')

    doc.render(context)

    output_path = output_dir / "generated_doc.docx"
    doc.save(output_path)

    os.system("start " + str(output_path))

    return 0


# # Press the green button in the gutter to run the script.
if __name__ == '__main__':
    open_home_menu()

    print("----------------------------------------")
    print("File generation completed. Check the OUTPUT folder.")
