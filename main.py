import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
import openpyxl


def load_data():
    path = r"C:\Users\kn849jw\PycharmProjects\Excel Viewer\list-countries-world.xlsx"
    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

    list_values = list(sheet.values)
    cols = list_values[0]
    tree = ttk.Treeview(root, column=cols, show="headings")
    for col_name in cols:
        tree.heading(col_name, text=col_name)
    tree.pack(expand=True, fill='y')

    for value_tuple in list_values[1:]:
        tree.insert('', tk.END, values=value_tuple)


root = ctk.CTk()
root.resizable(0, 0)
root.geometry('620x600')
root.title('Excel Viewer')

load_data()

root.mainloop()
