import zipfile

def main(agp0_submission_path: str) -> None:
    # Generate a report based on the contents of the provided ZIP file
    generate_agp_0_report(agp0_submission_path)

def generate_agp_0_report(agp0_submission_path: str) -> None:
    # List all files in the ZIP file
    file_names = list_files_in_zip(agp0_submission_path)
    # Organize the files by type (Word, Excel, PowerPoint, PDF)
    organized_files = organize_files_by_type(file_names)
    # Check the requirements for each type of file
    checked_requirements = check_all_requirements(organized_files)
    # Print the report to the command line
    print_report(checked_requirements)

def list_files_in_zip(zip_path: str) -> list:
    # Open the ZIP file and return the names of all files inside
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        return zip_ref.namelist()

def check_requirements(file_name: str, requirements: dict) -> None:
    # Check if the file name contains any of the keywords in requirements
    # If it does, add the file name to the corresponding key in the requirements dictionary
    file_name_lower = file_name.lower()
    for key, value in requirements.items():
        if key.lower() in file_name_lower:
            requirements[key].append(file_name)

def check_all_requirements(files_by_type: tuple) -> dict:
    # Initialize the requirements with empty lists
    requirements = {"SAS": [], "PAQ": [], "AR Log": [], "Decision Matrix": []}
    # Check each file against the requirements
    for file_type in files_by_type:
        for file_name in file_type:
            check_requirements(file_name, requirements)
    return requirements

def print_report(requirements: dict) -> None:
    # Build the report as a string and print it
    result = "SAS, PAQ, AR Log, and Decision Matrix are required.\n"
    for key, value in requirements.items():
        status = "exists: " + ', '.join(value) if value else "is missing!"
        result += f"{key} {status}\n"
    print(result)

def organize_files_by_type(file_names: list) -> tuple:
    # Sort the file names into different lists based on their extension
    word_file_names = []
    excel_file_names = []
    ppt_file_names = []
    pdf_file_names = []

    for file_name in file_names:
        extension = file_name.split('.')[-1].lower()
        if extension == "docx":
            word_file_names.append(file_name)
        elif extension == "xlsx":
            excel_file_names.append(file_name)
        elif extension == "pptx":
            ppt_file_names.append(file_name)
        elif extension == "pdf":
            pdf_file_names.append(file_name)

    return (word_file_names, excel_file_names, ppt_file_names, pdf_file_names)

if __name__ == "__main__":
    # Path to the ZIP file
    agp0_submission_path = "AGP0.zip"
    main(agp0_submission_path)
