# Student Records GUI Application

This program creates a graphical user interface (GUI) for entering student data into an Excel file. It utilizes the `openpyxl` module for reading and writing Excel files, and the `tkinter` module for creating the GUI.

**Note: This program was developed with the assistance of the OpenAI GPT-3 language model.**

## Features

- Provides a user-friendly interface for entering student data into an Excel file.
- Automatically creates a new `data.xlsx` file with headers if it does not exist.
- Requires the program to be run from the same directory as the `data.xlsx` file.
- Checks if the `data.xlsx` file is open by another user and prevents data entry until it is closed.
- Uses school options from a separate `schools.xlsx` file located in the same directory.
    - The `schools.xlsx` file should have a sheet named "Schools" with columns: "School Name", "Class 1", "Class 2", etc.
    - Each row in the `schools.xlsx` file represents a school with its corresponding classes.
- Utilizes nested dropdown menus for selecting schools and classes.
- Displays the newly added data in the data viewer.
- Supports double-clicking on a row in the viewer to edit the corresponding row data in the entry form.

## Limitations

The current version of the program has the following limitations:

- It does not validate the selection and validity of the school and class options.
- It does not validate the input and validity of the name and grade fields.
- It does not provide a feature to delete a row from the Excel file.
- It does not support sorting the data in the viewer by any column.
- It does not support filtering the data in the viewer by any column.
- It does not support searching the data in the viewer by any column.

## Prerequisites

Before running the program, ensure the following:

- Python 3.x is installed on your system.
- The necessary Python packages (`openpyxl`, `tkinter`) are installed.
- The `data.xlsx` file exists in the same directory as the program.
- The `schools.xlsx` file exists in the same directory as the program, with the required format.

## How to Run

To run the program, follow these steps:

1. Clone or download the program files to your local machine.
2. Open a terminal or command prompt and navigate to the program's directory.
3. Execute the following command to run the program:
    ```
    python write_student.py
    ```
4. The GUI application will open, allowing you to enter and view student data.

## Feedback and Contributions

Your feedback and contributions to this project are highly appreciated. If you encounter any issues, have suggestions for improvements, or would like to contribute to the project, please feel free to submit an issue or a pull request on the GitHub repository.

Thank you for using this Student Records GUI Application!
		