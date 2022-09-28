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

    open_home_menu() -> None
    generate_pre_agp_0_report() -> None
    generate_agp_0_report() -> None
    get_assessment_data() -> list
    get_similarity(rubric_text: str, rationale_text: str) -> float
    get_rubric_descriptions() -> list
    get_today_date_time() -> datetime
    get_risk_level(risk_score: int) -> str
    get_attribute_score(attribute_level: str) -> int
    fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None
    check_file_exist(file_name: str) -> None
"""
# this  function welcomes the user to the program
def open_home_menu() -> None:
    print("""
    _______
   /      /,
  / AIRE //
 /__V2__//
(______(/
    """)
    print("WELCOME TO OUR ARCHITECTURE INTAKE REVIEW ENGINE BOT!")
    input("Press enter to continue :)")

    invalid = True
    answer = 0

    while invalid:
        print("""
    __...--~~~~~-._   _.-~~~~~--...__
    //  1.Pre AGP 0  `V'   3.Exit      \\ 
   //  2.AGP 0        |                 \\ 
  //__...--~~~~~~-._  |  _.-~~~~~~--...__\\ 
 //__.....----~~~~._\ | /_.~~~~----.....__\\
====================\\|//====================
                dwb `---`
""")
        print("Please select which task you'd like to perform:\n")
        print("1. Pre AGP 0 Assessment: Use I&IT Decision Matrix to better assess a project item")
        print("2. AGP 0 Assessment: Ensure all mandatory files are submitted by Project group")
        print("3. Exit Program")
        answer = input()
        if 1 <= int(answer) <= 3:
            invalid = False

    if int(answer) == 1:
        generate_pre_agp_0_report()
    elif int(answer) == 2:
        print("--- In Development ---")
        input("Press enter to exit")
        exit()
        generate_agp_0_report()
    else:
        exit()

    return

def generate_pre_agp_0_report() -> None:

    initiative_name = input("What is the initiative name? (any name):")
    
    assessment_data = get_assessment_data()

    attribute_levels = assessment_data[0]
    risk_score = assessment_data[1]
    rationales = assessment_data[2]
    corporate_or_cluster = assessment_data[3]

    fill_in_template(initiative_name, attribute_levels, risk_score, rationales, corporate_or_cluster)

def generate_agp_0_report() -> None:
    print("1.Input a zip file of all artifact and supplement files (decision matrix, PAQ, SAS and etc)")
    print("2.Input a template file")
    print("3.Fill in the template with the files in the zip file")
    print("4.Save the filled template file")

def get_assessment_data() -> list:
    """
    Get attribute_levels, risk_score, rationales, and corporate_or_cluster from the assessment file
    """
    filename = input("Please enter the name of the assessment file (userinput.xlsx):")

    base_dir = Path(__file__).parent
    excel_path = base_dir / filename

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
    Get the text similarity between the rubric text and input text
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

# this function should take in the array of rubric descriptions and user input arrays
# based on the arrays, an appropriate set of results and conclusion should be reached
# report should open automatically (probably remind the users to save the generated report)
def fill_in_template(initiative_name: str, attribute_levels: list, risk_score: int, rationales: list, corporate_or_cluster: str) -> None:

    base_dir = Path(__file__).parent
    output_dir = base_dir / "Output"
    output_dir.mkdir(exist_ok=True)
    word_doc = base_dir / "Architecture Intake Review Engine Report Draft.docx"
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

    doc.render(context)

    check_file_exist('demo1.docx')

    output_path = output_dir / "generated_doc.docx"
    doc.save(output_path)

    os.system("start " + str(output_path))

# # Press the green button in the gutter to run the script.
if __name__ == '__main__':
    open_home_menu()

    print("----------------------------------------")
    print("""
        _________   _________
   ____/  AIRE   \ /  Report \____
 /| ------------- |  ------------ |\\
||| ------------- | ------------- |||
||| ------------- | ------------- |||
||| ------- ----- | ------------- |||
||| ------------- | ------------- |||
||| ------------- | ------------- |||
|||  ------------ | ----------    |||
||| ------------- |  ------------ |||
||| ------------- | ------------- |||
||| ------------- | ------ -----  |||
||| ------------  | ------------- |||
|||_____________  |  _____________|||
L/_____/--------\\_//W-------\_____\J
    """)
    print("File generation completed. Check the OUTPUT folder.")
