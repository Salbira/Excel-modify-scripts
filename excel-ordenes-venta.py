import openpyxl

def agregar_filas_extra(archivo_entrada, archivo_salida, nombre_hoja):
    # Cargar el archivo Excel
    wb = openpyxl.load_workbook(archivo_entrada)
    
    # Seleccionar la hoja específica por su nombre
    sheet = wb[nombre_hoja]

    # Crear una nueva hoja para el archivo de salida
    wb_salida = openpyxl.Workbook()
    sheet_salida = wb_salida.active
    
    # Definir la columna de "nombre de tienda"
    columna_nombre_tienda = None

    # Buscar la columna "nombre de tienda" y agregar filas adicionales debajo de cada celda
    for col_idx, col in enumerate(sheet.iter_cols(), start=1):
        for cell in col:
            if cell.value == "Nombre de la tienda":
                columna_nombre_tienda = col_idx
                break
        if columna_nombre_tienda:
            break

    if columna_nombre_tienda:
        for row_idx, row in enumerate(sheet.iter_rows(), start=1):
            nombre_tienda = row[columna_nombre_tienda - 1].value
            sheet_salida.append([cell.value for cell in row])
            if nombre_tienda == "Nombre de la tienda":
                continue
            for i in range(3):
                sheet_salida.append([''] * sheet.max_column)

    # Guardar el archivo de salida
    wb_salida.save(archivo_salida)


archivo_entrada = 'Tienda.xlsx'
archivo_salida = 'archivo_salida.xlsx'
nombre_hoja = 'Sheet1'
agregar_filas_extra(archivo_entrada, archivo_salida, nombre_hoja)
