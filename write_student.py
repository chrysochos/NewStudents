import tkinter as tk
from display_students import Display_students
from student_entry import StudentEntry

'''
This program creates a GUI to enter student data into an Excel file.
The program uses the openpyxl module to read and write Excel files.
The program uses the tkinter module to create the GUI.
Up to now the program:
- Can not create a new data.xlsx file. We must create it manually.
- Must be run from the same directory as the data.xlsx file.
- Does not check if the data.xlsx file is open by another user.
- Does not check if the data.xlsx file exists.
- Does not check if the data.xlsx file is empty and does not have column headers. 
- It should write the column headers if they are missing.
- Does not check if the data.xlsx file has the correct column headers.\
- Does not check if the data.xlsx file has the correct number of columns.
- The menu options are given by the schools.xlsx file. The schools.xlsx file must be in the same directory as the script.
- The schools.xlsx file must have a sheet named "Schools".
- The schools.xlsx file must have a column named "School Name".
- The schools.xlsx file must have a column named "Class 1", "Class 2", "Class 3", etc.
- The schools.xlsx file must have a row for each school.
- The schools.xlsx file must have a row for each class.
- When the school is changed, the class dropdown menu is not updated.
- The program does not check if the school and class are valid.
- The program does not check if the school and class are selected.
- The program does not check if the name and grade are entered.
- The program does not check if the name and grade are valid.

- The program should show the newly added data in the data viewer.
- The program should allow the user to double-click a row in the viewer and the row data to go for editing in the entry form
- The program should allow the user to edit the data in the entry form and save the changes to the Excel file.
- The program should allow the user to delete a row from the Excel file.
- The program should allow the user to sort the data in the viewer by any column.
- The program should allow the user to filter the data in the viewer by any column.
- The program should allow the user to search the data in the viewer by any column.
- The program should allow the user to export the data in the viewer to a CSV file.
- The program should allow the user to import data from a CSV file to the Excel file.
- The program should allow the user to create a new Excel file.
- The program should allow the user to create a new Excel file from a CSV file.


'''

    
if __name__ == "__main__":
    root = tk.Tk()

    display_students = Display_students(master=root)

    student_entry = StudentEntry(master=root)
 
    student_entry.inject_display_students = display_students
    display_students.inject_student_entry = student_entry
    
    student_entry.mainloop()
    display_students.mainloop()