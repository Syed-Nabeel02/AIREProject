import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os
from pathlib import Path
import openpyxl
from difflib import SequenceMatcher


def menu():
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
        ini_name = input("What is the initiative name?:")
        user_input(ini_name)
    elif int(answer) == 2:
        print("This version has not been released yet!")
        input("Press enter to exit")
        exit()
    else:
        exit()
    return


def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()


# this  function welcomes the user to the program
def welcome():
    print("WELCOME TO OUR ARCHITECTURE INTAKE REVIEW ENGINE BOT!")
    input("Press enter to continue :)")


# this function contains the criteria
# this function should load in the data from Rubric
def criteria():
    base_dir = Path(__file__).parent
    # word_doc = base_dir / "Architecture Intake Review Engine Report Draft.docx"
    output_dir = base_dir / "Output"
    output_dir.mkdir(exist_ok=True)
    # doc = DocxTemplate(word_doc)
    excel_path = base_dir / "IIT-EA-Decision-Matrix.xlsx"
    df = pd.read_excel(excel_path, sheet_name="Rubric")
    options = []
    for record in df.to_dict(orient="records"):
        c = record['Description']
        options.append(c)
    return options


def user_input(name):
    base_dir = Path(__file__).parent
    filename = input("Please enter the name of your file:")
    excel_path = base_dir / filename
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = wb['Matrix']
    crit_assessment = [sheet['D10'].value, sheet['D11'].value, sheet['D12'].value, sheet['D13'].value,
                       sheet['D14'].value]
    score = sheet['D15'].value
    rational = [sheet['C10'].value, sheet['C11'].value, sheet['C12'].value, sheet['C13'].value,
                sheet['C14'].value]
    corporate_cluster = sheet['D22'].value
    report(crit_assessment, score, rational, name, corporate_cluster)
    return 0


def get_date():
    today = datetime.now()
    date_time = today.strftime("%m/%d/%Y, %H:%M:%S")
    return date_time


def risk_assessment(score):
    if score < 7:
        risk = 'Low'
    elif 7 <= score < 11:
        risk = 'Medium'
    else:
        risk = 'High'
    return risk


def evaluate(rank):
    if rank == 'Low':
        return 0
    elif rank == 'Medium':
        return 1
    else:
        return 2


# this function should take in the array of criteria and user input arrays
# based on the arrays, an appropriate set of results and conclusion should be reached
# report should open automatically (probably remind the users to save the generated report)
def report(crit_assessment, score, rational, name, corporate_cluster):
    base_dir = Path(__file__).parent
    word_doc = base_dir / "Architecture Intake Review Engine Report Draft.docx"
    output_dir = base_dir / "Output"
    output_dir.mkdir(exist_ok=True)
    doc = DocxTemplate(word_doc)
    if score < 7:
        gov = 'does not'
    else:
        gov = 'does'
    options = criteria()
    business_scope = [options[0], options[1], options[2]]
    it_solution = [options[3], options[4], options[5]]
    technology_up = [options[6], options[7], options[8]]
    info_req = [options[9], options[10], options[11]]
    info_sens = [options[12], options[13], options[14]]
    context = {'date': get_date(),
               'initiative': name,
               'score': score,
               'risk': risk_assessment(score),
               'gov': gov,
               'ca': crit_assessment[0],
               'ca1': crit_assessment[1],
               'ca2': crit_assessment[2],
               'ca3': crit_assessment[3],
               'ca4': crit_assessment[4],
               'rational': rational[0],
               'rational1': rational[1],
               'rational2': rational[2],
               'rational3': rational[3],
               'rational4': rational[4],
               'comp': business_scope[evaluate(crit_assessment[0])],
               'comp1': it_solution[evaluate(crit_assessment[1])],
               'comp2': technology_up[evaluate(crit_assessment[2])],
               'comp3': info_req[evaluate(crit_assessment[3])],
               'comp4': info_sens[evaluate(crit_assessment[4])],
               'cluster_corporate': corporate_cluster,
               's': str(round((similar(business_scope[evaluate(crit_assessment[0])], rational[0]) * 100), 2)) + '%',
               's1': str(round((similar(it_solution[evaluate(crit_assessment[1])], rational[1]) * 100), 2)) + '%',
               's2': str(round((similar(technology_up[evaluate(crit_assessment[2])], rational[2]) * 100), 2)) + '%',
               's3': str(round((similar(info_req[evaluate(crit_assessment[3])], rational[3]) * 100), 2)) + '%',
               's4': str(round((similar(info_sens[evaluate(crit_assessment[4])], rational[4]) * 100), 2)) + '%',
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
    welcome()
    menu()
    print("done")
