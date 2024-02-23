from openpyxl import Workbook, load_workbook
from utils import clear_entries, update_table
import tkinter as tk

def add_data(column_headers, entry_fields, table_text):
    data = [entry_fields[header].get() for header in column_headers]

    try:
        workbook = load_workbook('data.xlsx')
        sheet = workbook.active
    except FileNotFoundError:
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(column_headers)

    sheet.append(data)
    workbook.save('data.xlsx')

    # Update the table display
    update_table(table_text)

    # Clear the entry fields
    clear_entries(entry_fields)


    # table_text.config(state=tk.DISABLED)
