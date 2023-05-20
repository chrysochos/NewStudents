import openpyxl
import tkinter as tk
# import pandas as pd
# import ipywidgets as widgets
# from IPython.display import display
# from tkinter import messagebox
import tkinter.ttk as ttk
from openpyxl import Workbook

class Display_students(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        #self.master.title("Excel Data Viewer")
        self.tree = ttk.Treeview(self.master)
        self.grid()
        self.re_edit = False

        # Create a canvas widget with a scrollbar
        self.canvas = tk.Canvas(self, width=500, height=300)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.config(yscrollcommand=self.scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.config(scrollregion=self.canvas.bbox("all")))

        # Create a frame inside the canvas to hold the data
        self.data_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0,0), window=self.data_frame, anchor=tk.NW)

        try:
            # Load the data from the Excel file
            self.wb = openpyxl.load_workbook("data.xlsx")
            self.ws = self.wb.active
        except FileNotFoundError:
            # If the file doesn't exist, create a new workbook and add the headers
            wb = Workbook()
            sheet = wb.active

            # Add the headers
            headers = ["Name", "Grade", "Gender", "OldSchool", "OldSchoolClass"]
            sheet.append(headers)

            # Save the workbook with the headers
        #     wb.save("data.xlsx")
        # self.ws = self.wb.active

        # Create the treeview to display the data
        self.treeview = ttk.Treeview(self.data_frame, selectmode='browse')
        self.treeview.grid(row=1, column=0, padx=5, pady=5)

        # Define the columns
        self.treeview['columns'] = ('name', 'grade', 'gender', 'school', 'class')
        self.treeview.heading('#0', text='ID')
        self.treeview.column('#0', width=0, stretch='no')
        self.treeview.heading('name', text='Name')
        self.treeview.column('name', width=150, anchor='w')
        self.treeview.heading('grade', text='Grade')
        self.treeview.column('grade', width=75, anchor='center')
        self.treeview.heading('gender', text='Gender')
        self.treeview.column('gender', width=75, anchor='center')
        self.treeview.heading('school', text='School')
        self.treeview.column('school', width=150, anchor='w')
        self.treeview.heading('class', text='Class')
        self.treeview.column('class', width=75, anchor='center')


        # Display the data rows
        row_num = 1
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            self.treeview.insert('', 'end', text=row_num, values=row)
            row_num += 1

        # Bind the double-click event to the on_row_select method
        self.treeview.bind('<Double-1>', self.on_row_select)

    def on_row_select(self, event):
        # Get the selected row
        print("Hi I'm in on_row_select of display_students.py")
        print(event)
        self.re_edit = True
        self.selected_item = self.treeview.selection()[0]
        values = self.treeview.item(self.selected_item, 'values')

      
        # Fill the entry form with the selected row data
        self.inject_student_entry.name_entry.delete(0, 'end')
        self.inject_student_entry.name_entry.insert(0, values[0])
        self.inject_student_entry.grade_entry.delete(0, 'end')
        self.inject_student_entry.grade_entry.insert(0, values[1])
        self.inject_student_entry.gender_var.set(values[2])
        self.inject_student_entry.school_var.set(values[3])
        self.inject_student_entry.class_var.set(values[4])


    def refresh_student_display(self):
        # Display the data rows
        print("Hi I'm in refresh_student_display of display_students.py")
        self.treeview.delete(*self.treeview.get_children())
        row_num = 1
        for row in self.ws.iter_rows(min_row=2, values_only=True):
            self.treeview.insert('', 'end', text=row_num, values=row)
            row_num += 1

    def replace_student(self,values):
        selected_item = self.treeview.selection()[0]
        print("Hi I'm in replace_student of display_students.py")
        print(selected_item)
        #open the workbook data.xlsx
        wb = openpyxl.load_workbook("data.xlsx")
        # find the row in data.xlsx that corresponds to the selected row
        ws = wb.active



def main():
    root = tk.Tk()
    # Local main for testing Object1
    display_students = Display_students(master=root)
    # Test Object1 functionality
    display_students.mainloop()

if __name__ == "__main__":
    main()