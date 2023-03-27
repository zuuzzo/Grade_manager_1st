import pandas as pd
import tkinter as tk
from tkinter import messagebox as mb

def delete_student():
    name = name_var.get()
    try:
        df = pd.read_excel('information.xlsx')
        df = df[df['Name'] != name]
        df.to_excel('information.xlsx', index=False)
        mb.showinfo("Success", "Student deleted successfully")
    except:
        mb.showerror("Error", "Student not found")

def add_student():
    name = name_var.get()
    grade = grade_var.get()

    if not grade.isdigit():
        mb.showerror("Error", "Grade must be a number")
        return

    grade = int(grade)
    if grade < 9 or grade > 11:
        mb.showerror("Error", "Grade must be between 9 and 11")
        return

    new_data = {'Name': [name], 'Grade': [grade]}
    df = pd.DataFrame(new_data)
    try:
        df_existing = pd.read_excel('information.xlsx')
        df_existing = pd.concat([df_existing, df], ignore_index=True)
        df_existing.to_excel('information.xlsx', index=False)
        mb.showinfo("Success", "Student added successfully")
    except:
        df.to_excel('information.xlsx', index=False)
        mb.showinfo("Success", "Student added successfully")

def display_information():
    try:
        df = pd.read_excel('information.xlsx')
        mb.showinfo("Information", df.to_string(index=False))
    except:
        mb.showerror("Error", "No information found")

window = tk.Tk()
window.geometry("500x200")
window.title("Grade Manager")

name_label = tk.Label(window, text="Name:")
name_label.grid(row=0, column=0, padx=5, pady=5)

name_var = tk.StringVar()
name_entry = tk.Entry(window, textvariable=name_var)
name_entry.grid(row=0, column=1, padx=5, pady=5)

grade_label = tk.Label(window, text="Grade:")
grade_label.grid(row=1, column=0, padx=5, pady=5)

grade_var = tk.StringVar()
grade_entry = tk.Entry(window, textvariable=grade_var)
grade_entry.grid(row=1, column=1, padx=5, pady=5)

add_button = tk.Button(window, text="Add Student", command=add_student)
add_button.grid(row=2, column=0, padx=5, pady=5)

delete_button = tk.Button(window, text="Delete Student", command=delete_student)
delete_button.grid(row=2, column=1, padx=5, pady=5)

display_button = tk.Button(window, text="Display Information", command=display_information)
display_button.grid(row=3, column=0, padx=5, pady=5)

exit_button = tk.Button(window, text="Exit", command=window.destroy)
exit_button.grid(row=3, column=1, padx=5, pady=5)

window.mainloop()