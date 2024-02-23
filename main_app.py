import tkinter as tk
from tkinter import ttk
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

    for i in range(5):
        if column_headers[i] != 'Depto_solicitante':
           label = tk.Label(frame1, text=column_headers[i])
           label.grid(row=0, column=i, padx=5, pady=2)

           entry = tk.Entry(frame1)
           entry.grid(row=1, column=i, padx=5, pady=2)
           entry_fields[column_headers[i]] = entry

    frame2 = tk.Frame(root)
    frame2.pack()

    # Exclude 'Depto_solicitante' from UI
    for i in range(5, 10):
        if column_headers[i] not in ['Tipo_acceso', 'Ejecutando'] :
            label = tk.Label(frame2, text=column_headers[i])
            label.grid(row=0, column=i-5, padx=5, pady=2)

            entry = tk.Entry(frame2)
            entry.grid(row=1, column=i-5, padx=5, pady=2)
            entry_fields[column_headers[i]] = entry


    departamentos_label = tk.Label(root, text="Departamentos")
    departamentos_label.pack(padx=5, pady=2)

    #Select Departamentos
    departamentos_combobox = ttk.Combobox(root, values=["Abastecimiento", "Dirección", "hidrografia", "Instrucción",
                                          "Oceanografiía", "Producción", "Servicios a Terceros", "Servicios Generales", "Tecnologías de Información", "Terreno"])
    departamentos_combobox.pack(padx=5, pady=2)
    
    # Select tipo acceso
    tipo_acceso_label = tk.Label(root, text="Tipo_acceso")
    tipo_acceso_label.pack(padx=5, pady=2)
    tipo_acceso_combobox = ttk.Combobox(root, values=["L", "LyE"])
    tipo_acceso_combobox.pack(padx=5, pady=2)

    # Select Ejecutando
    ejecutando_label = tk.Label(root, text="Ejecutando")
    ejecutando_label.pack(padx=5, pady=2)
    ejecutando_combobox = ttk.Combobox(root, values=["SI", "NO"])
    ejecutando_combobox.pack(padx=5, pady=2)

    add_button = tk.Button(root, text='Añadir Datos', command=lambda: add_data(
    column_headers, entry_fields, departamentos_combobox, tipo_acceso_combobox, ejecutando_combobox, table_text))
    add_button.pack(pady=10)

    # open_button = tk.Button(root, text='Abrir Carpeta',
    #                         command=open_excel_file)
    # open_button.pack(pady=10, padx=100)

    open_search_button = tk.Button(
        root, text='Buscar Tipo_solicitud', command=lambda: create_search_window(root))
    open_search_button.pack(pady=10)

    row_number_entry = tk.Entry(root)
    row_number_entry.pack(pady=2)

    delete_button = tk.Button(root, text='Eliminar filas seleccionadas', command=lambda: delete_selected_rows(
        table_text, row_number_entry, row_display_label))
    delete_button.pack(pady=10)

    row_display_label = tk.Label(root, text='', padx=10)
    row_display_label.pack(side=tk.BOTTOM, pady=5)

    table_text = tk.Text(root, height=10, width=200,
                         wrap=tk.WORD, state=tk.DISABLED)
    table_text.pack(pady=10)
    # Initial table display
    update_table(table_text)

    # Start the main loop
    root.resizable(True, True)
    root.mainloop()


if __name__ == "__main__":
    main()
