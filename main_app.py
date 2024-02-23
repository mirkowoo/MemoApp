import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from excel_operations import add_data
from utils import delete_selected_rows, update_table, create_search_window


def main():
    root = tk.Tk()
    root.title('Excel Data Manager - Mirko Franichevic')

    # Modify the column_headers list as needed
    column_headers = ['Numero_memo', 'Tipo_solicitud', 'Fecha_memo', 'Depto_solicitante',
                      'Resp_memo', 'RUT', 'Usuario', 'Tipo_acceso', 'Ejecutando', 'Ingresado_por']

    entry_fields = {}

    frame1 = tk.Frame(root)
    frame1.pack()



    for i in range(4):
        if column_headers[i] not in ['Depto_solicitante', 'Tipo_acceso', 'Ejecutando']:
            label = tk.Label(frame1, text=column_headers[i])
            label.grid(row=0, column=i, padx=5, pady=2)

            if column_headers[i] == 'Fecha_memo':
                date_entry = DateEntry(frame1, width=12, background='darkblue', foreground='white', borderwidth=2,date_pattern='dd/mm/yyyy')
                date_entry.grid(row=1, column=i, padx=5, pady=2)
                entry_fields[column_headers[i]] = date_entry
            else:
                entry = tk.Entry(frame1)
                entry.grid(row=1, column=i, padx=5, pady=2)
                entry_fields[column_headers[i]] = entry

    row_offset = 2 

    for i in range(4, 10):
        if column_headers[i] not in ['Tipo_acceso', 'Ejecutando']:
            label = tk.Label(frame1, text=column_headers[i])
            label.grid(row=row_offset, column=i - 4, padx=5, pady=2)

            entry = tk.Entry(frame1)
            entry.grid(row=row_offset + 1, column=i - 4, padx=5, pady=2)
            entry_fields[column_headers[i]] = entry

    # Display 'Tipo_acceso', 'Departamentos', 'Ejecutando' horizontally
    labels_comboboxes = [('Tipo_acceso', ["L", "LyE"]),
                         ('Departamentos', ["Abastecimiento", "Dirección", "hidrografia", "Instrucción",
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

    add_button = tk.Button(root, text='Añadir Datos', command=lambda: add_data(
        column_headers, entry_fields, departamentos_combobox, tipo_acceso_combobox, ejecutando_combobox, table_text))
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

    table_text = tk.Text(root, height=10, width=200, wrap=tk.WORD, state=tk.DISABLED, borderwidth=1, relief=tk.SOLID)
    table_text.pack(pady=10)

    # Initial table display
    update_table(table_text)

    # Start the main loop
    root.resizable(True, True)
    root.mainloop()


if __name__ == "__main__":
    main()
