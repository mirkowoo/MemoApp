import tkinter as tk
from openpyxl import load_workbook

def search_data(root, username_entry, result_label):
    username_to_search = username_entry.get()

    try:
        workbook = load_workbook('data.xlsx')
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[6] == username_to_search:
                result_label.config(text=f"Tipo_solicitud: {row[7]}")
                return

        result_label.config(text=f"Usuario '{username_to_search}' no encontrado.")

    except FileNotFoundError:
        result_label.config(text="Error: File not found.")

def open_search_window(root):
    search_window = tk.Toplevel(root)
    search_window.title("Buscar Usuario")

    username_entry = tk.Entry(search_window)
    username_entry.grid(row=0, column=1, padx=5, pady=5)

    search_button = tk.Button(search_window, text="Buscar", command=lambda: search_data(root, username_entry, result_label))
    search_button.grid(row=0, column=2, padx=5, pady=5)

    result_label = tk.Label(search_window, text="")
    result_label.grid(row=1, column=0, columnspan=3, padx=5, pady=5)

