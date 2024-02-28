import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from excel_operations import add_data
from utils import delete_selected_rows, update_table, create_search_window
from babel import numbers


def main():
    appMemo()
    # menu_screen()


def menu_screen():
    root = tk.Tk()
    root.title('Menu Screen')

    program_button = tk.Button(root, text='Go to Program', command=appMemo)
    program_button.pack(pady=10)

    blank_button = tk.Button(
        root, text='Go to Blank Screen', command=blank_screen)
    blank_button.pack(pady=10)

    root.mainloop()


def blank_screen():
    # Create the blank screen window
    blank_window = tk.Toplevel()
    blank_window.title('Custom Table')

    # Function to add a new entry field for a column name
    def add_column():
        new_entry = tk.Entry(columns_frame)
        new_entry.grid(row=len(column_entries) + 1, column=0, padx=5, pady=2)
        column_entries.append(new_entry)

    def send_columns(array):
        a = 0

    # Frame to contain the column entry fields
    columns_frame = tk.Frame(blank_window)
    columns_frame.pack(padx=10, pady=10)

    # Button to add a new column entry field
    add_column_button = tk.Button(
        columns_frame, text='Añadir Columna', command=add_column)
    add_column_button.grid(row=0, column=0, pady=10)

    # List to keep track of column entry fields
    column_entries = []

    # Button to apply formatting to the columns
    apply_format_button = tk.Button(
        blank_window, text='Añadir Formato', command=lambda: apply_format(column_entries))
    apply_format_button.pack(pady=10)

    send_data_button = tk.Button(
        blank_window, text='Enviar', command=lambda: send_columns(column_entries))
    send_data_button.pack(pady=10)
    # Function to apply formatting to the columns

    def apply_format(entries):
        # Iterate over the column entry fields and print their values
        for entry in entries:
            print(entry.get())

    # Run the blank screen window
    blank_window.mainloop()


def appMemo():
    root = tk.Tk()
    root.title('Excel Data Manager - Mirko Franichevic')

    # Modify the column_headers list as needed
    column_headers = ['Numero_memo', 'Tipo_solicitud', 'Fecha_memo', 'Servidor', 'Carpeta', 'Depto_solicitante',
                      'Resp_memo', 'RUT', 'Usuario', 'Tipo_acceso', 'Ejecutando', 'Ingresado_por']

    entry_fields = {}

    frame1 = tk.Frame(root)
    frame1.pack()
    saltados = 0
    # display
    for i in range(6):
        if column_headers[i] not in ['Depto_solicitante', 'Tipo_acceso', 'Ejecutando', 'Tipo_solicitud']:
            label = tk.Label(frame1, text=column_headers[i])
            label.grid(row=0, column=i - saltados, padx=5, pady=2)

            if column_headers[i] == 'Fecha_memo':
                date_entry = DateEntry(frame1, width=12, background='black',
                                       foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                date_entry.grid(row=1, column=i - saltados, padx=5, pady=2)
                entry_fields[column_headers[i]] = date_entry
            else:
                entry = tk.Entry(frame1)
                entry.grid(row=1, column=i-saltados, padx=5, pady=2)
                entry_fields[column_headers[i]] = entry
        else:
            saltados += 1

    row_offset = 2
    saltados = 0
    for i in range(6, 12):
        if column_headers[i] not in ['Tipo_acceso', 'Ejecutando', 'Tipo_solicitud']:
            label = tk.Label(frame1, text=column_headers[i])
            label.grid(row=row_offset, column=i - 6 - saltados, padx=5, pady=2)

            entry = tk.Entry(frame1)
            entry.grid(row=row_offset + 1, column=i - 6 - saltados, padx=5, pady=2)
            entry_fields[column_headers[i]] = entry
        else:
            saltados += 1

    # selects -> comboboxes
    labels_comboboxes = [
        ('Tipo_solicitud', ["MEMO", "CORREO"]),
        ('Tipo_acceso', ["L", "LyE"]),
        ('Departamentos', ["Abastecimiento", "Dirección", "Hidrografia", "Instrucción",
                           "Oceanografía", "Producción", "Servicios a Terceros", "Servicios Generales", "Tecnologías de Información", "Terreno"]),
        ('Ejecutando', ["SI", "NO"])]

    for i, (label_text, values) in enumerate(labels_comboboxes):
        label = tk.Label(frame1, text=label_text)
        label.grid(row=row_offset + 2, column=i, padx=5, pady=2)

        combobox = ttk.Combobox(frame1, values=values)
        combobox.grid(row=row_offset + 3, column=i, padx=5, pady=2)
        entry_fields[label_text] = combobox

    # Move these comboboxes creation here
    departamentos_combobox = entry_fields['Departamentos']
    tipo_acceso_combobox = entry_fields['Tipo_acceso']
    ejecutando_combobox = entry_fields['Ejecutando']
    tipo_solicitud_combobox = entry_fields['Tipo_solicitud']

    add_button = tk.Button(root, text='Añadir Datos', command=lambda: add_data(
        column_headers, entry_fields, tipo_solicitud_combobox, departamentos_combobox, tipo_acceso_combobox, ejecutando_combobox, table_text,row_display_label))
    add_button.pack(pady=10)

    # open_button = tk.Button(root, text='Abrir Carpeta',
    #                         command=open_excel_file)
    # open_button.pack(pady=10, padx=100)

    open_search_button = tk.Button(
        root, text='Buscar Tipo_acceso', command=lambda: create_search_window(root))
    open_search_button.pack(pady=10)

    row_number_entry = tk.Entry(root)
    row_number_entry.pack(pady=2)

    delete_button = tk.Button(root, text='Eliminar filas seleccionadas', command=lambda: delete_selected_rows(
        table_text, row_number_entry, row_display_label))
    delete_button.pack(pady=10)

    row_display_label = tk.Label(root, text='', padx=10)
    row_display_label.pack(side=tk.BOTTOM, pady=5)

    table_text = tk.Text(root, height=10, width=200, wrap=tk.WORD,
                         state=tk.DISABLED, borderwidth=1, relief=tk.SOLID)
    table_text.pack(pady=10)

    # go_back_button = tk.Button(root, text='Go Back', command=root.destroy)
    # go_back_button.pack(pady=10)

    # Initial table display
    update_table(table_text)

    # Start the main loop
    root.resizable(True, True)
    root.mainloop()


if __name__ == "__main__":
    main()
