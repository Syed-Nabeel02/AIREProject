# Importing modules corresponding to different functionalities
import generate_intake_forms
import compare_operational_plans
import generate_risk_assessment_report

def main_menu():
    # Print the main menu of the Architecture Intake Review Engine (AIRE)
    print("AIRE - Architecture Intake Review Engine\n")
    print("Choose one of the following options:")
    print("1. Generate Intake Forms")
    print("2. Compare Operational Plans")
    print("3. Generate a Risk Assessment Report")
    print("4. Exit")

    # Prompt the user to enter their choice
    choice = input("\nEnter your choice (1-4): ")

    # Execute the corresponding functionality based on the user's choice
    if choice == '1':
        generate_intake_forms.main()  # Call the main function of the intake forms module
    elif choice == '2':
        compare_operational_plans.main()  # Call the main function of the operational plans comparison module
    elif choice == '3':
        generate_risk_assessment_report.main()  # Call the main function of the risk assessment report module
    elif choice == '4':
        print("Exiting the program.")  # Exit the program
        exit()
    else:
        print("Invalid choice. Please enter a number between 1 and 4.")  # Print an error message for an invalid choice
        main_menu()  # Recursively call the main_menu function to prompt the user again

if __name__ == "__main__":
    main_menu()  # Start the main menu when the script is run directly
