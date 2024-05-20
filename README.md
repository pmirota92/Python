# Grade Book Application - README

Welcome to the **Grade Book Application**, a Python-based tool designed to manage and analyze student grades efficiently. This application provides various functionalities to calculate averages and generate rankings for subjects and students. Here’s a step-by-step guide to help you navigate and utilize the Grade Book Application effectively.

## Grade Book App Demo
![ Python-Dziennik-Ocen](demo/GradeBook.gif)

## Features

1. **Calculate Average Grade for a Subject**
2. **Calculate Average Grade for a Student**
3. **Calculate Overall Average Grades**
4. **Generate Subject-wise Student Rankings**
5. **Generate Overall Student Rankings**

## Installation

To run this application, ensure you have the following dependencies installed:
- Python 3.x
- `termcolor` module for colored terminal output
- `openpyxl` module for Excel file operations

Install the required Python packages using pip:
```bash
pip install termcolor openpyxl
```

## Setup

1. **Excel File Preparation:**
   Ensure your grades are stored in an Excel file (`oceny.xlsx`) located at:
   ```
   C:\Users\...\oceny.xlsx
   ```
   The file should have separate sheets for each subject, with the first column for student names and the second column for their grades.

## Usage

Run the application by executing the script. This will launch the main menu:
```bash
python gradebook.py
```

### Main Menu Options

1. **Calculate Average Grade for a Subject:**
   - Select option `[1]`.
   - Enter the subject name.
   - The application will display the average grade for the specified subject.

2. **Calculate Average Grade for a Student:**
   - Select option `[2]`.
   - Enter the student’s name.
   - The application will display the average grade for the specified student across all subjects.

3. **Calculate Overall Average Grades:**
   - Select option `[3]`.
   - The application will display the average grade for all students across all subjects.

4. **Generate Subject-wise Student Rankings:**
   - Select option `[4]`.
   - Enter the subject name.
   - The application will display a ranking of students for the specified subject.

5. **Generate Overall Student Rankings:**
   - Select option `[5]`.
   - The application will display a ranking of students based on their average grades across all subjects.

6. **Exit the Application:**
   - Select option `[6]` to exit the program.

### Example Workflow

1. Launch the main menu.
2. Choose an option from the menu (e.g., `[1]` for subject average).
3. Follow the prompts to input the required data (e.g., subject name or student name).
4. View the results directly in the terminal.
5. After each operation, choose whether to exit or return to the main menu.

## Contribution

If you encounter any issues or have suggestions for improvements, please feel free to contribute to this project by creating a pull request or reporting an issue on the project's GitHub repository.

---

Enjoy using the Grade Book Application! If you have any questions or need further assistance, please refer to the code comments for detailed explanations of each function. Happy grading!
