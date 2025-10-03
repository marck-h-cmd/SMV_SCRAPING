import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries


# Color de relleno de encabezado (Azul Oscuro/Teal)
HEADER_FILL = PatternFill(start_color="003366", end_color="003366", fill_type="solid")

# Color de relleno de encabezado (Azul Oscuro/Teal)
ENCABEZADO_NARANJA = PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid")

ENCABEZADO_PURPURA = PatternFill(start_color="BE78F8", end_color="BE78F8", fill_type="solid")

ENCABEZADO_CELESTE = PatternFill(start_color="65FFD7", end_color="65FFD7", fill_type="solid")

# Color de relleno para la sección principal (Gris Claro)
SECTION_FILL = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

# Fuente para encabezados (Blanco y Negrita)
HEADER_FONT = Font(size=9, bold=True, color="FFFFFF")
Contenido = Font(name="Franklin Gothic Medium Cond", size=9)

# Estilo de borde (borde fino para todas las celdas de datos)
THIN_SIDE = Side(style='thin', color="000000")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

fuente_titulo = Font(size=13.5, bold=True)
negrita = Font(bold=True)
cuentas = Alignment(horizontal='center', vertical='center')

# Formato de SIMBOLO DECIMAL (.), MILES (,) SIMBOLO NEGATIVO (-)
FORMATO_NUMERICO_FINANCIERO = '#,##0.00;-#,##0.00'
FORMATO_NUMERICO = '#,##0;-#,##0.'
FORMATO_PORCENTAJE_DOS_DECIMALES = '0.00%'

# -----------------------------
def formato_xls_xlsx(CARPETA_FINANCIEROS):
    #FORMATO DE LOS ESTADOS FINANCIEROS - Ultimos 5 años
    for archivo in os.listdir(CARPETA_FINANCIEROS):
        if archivo.endswith(".xls"):
            path_xls = os.path.join(CARPETA_FINANCIEROS, archivo)
            path_xlsx = path_xls + "x"
            dir_path = os.path.dirname(path_xls)    #Ruta del excel
            nombre = os.path.basename(dir_path)    #Nombre de la empresa
            nombreEmpresa = nombre.replace('_', ' ')
            tablas = pd.read_html(path_xls)

            # 1. GUARDAR DATAFRAME
            with pd.ExcelWriter(path_xlsx, engine="openpyxl") as writer:
                for i, df in enumerate(tablas, start=1):

                    if df.shape[1] >= 4 and i!=3:   #TODOS MENOS LA DE PATRIMONIO O FIRMANTES ELIMINAN LA COLUMNA B
                        df_cleaned = df.drop(df.columns[[1, 3]], axis=1, errors='ignore')
                    else:
                        df_cleaned = df

                    def clean_and_coerce(series):
                        s = series.astype(str)
                        s = s.str.replace('(', '-', regex=False).str.replace(')', '', regex=False)
                        s = s.str.replace(',', '', regex=False)
                        return pd.to_numeric(s, errors='coerce')

                    if df_cleaned.shape[1] >= 3 and (i!=3 and i!=6):
                        df_cleaned.iloc[:, [1, 2]] = df_cleaned.iloc[:, [1, 2]].apply(clean_and_coerce)
                    else:
                        df_cleaned.iloc[:, 2:df_cleaned.shape[1]] = df_cleaned.iloc[:, 2:df_cleaned.shape[1]].apply(clean_and_coerce)

                    df_cleaned.to_excel(
                        writer, 
                        sheet_name=f"Hoja{i}", 
                        index=False,
                        header=True,
                        startrow=6,
                        startcol=2
                    )

            # 2. Declaración de hojas de ESTADOS:
            wb = load_workbook(path_xlsx)
            situaFinanciera = wb['Hoja1']
            estaresultados = wb['Hoja2']
            patrimonio = wb['Hoja3']
            flujoEfectivo = wb['Hoja4']

            # 3. ABRIR CON OPENPYXL PARA APLICAR ESTILOS MANUALES
            FormatoSituacionFinanciera(situaFinanciera, nombreEmpresa)
            FormatoResultados(estaresultados, nombreEmpresa)
            FormatoPatrimonio(wb, patrimonio, nombreEmpresa)
            FormatoFlujoEfectivo(flujoEfectivo, nombreEmpresa)

            # 4. ELIMINAR HOJAS QUE NO SON NECESARIAS
            wb.remove(wb['Hoja5'])
            wb.remove(wb['Hoja6'])

            # 5. Guardar el archivo con estilos
            wb.save(path_xlsx)

def hojas(ws):
    if ws.title == 'Hoja1':
        ws['A1'].value = "ESTADO DE SITUACIÓN FINANCIERA"
        FILAS_GRISES = [8,9,24,25,43,44,45,46,59,60,72,73,74,82,83]

    elif ws.title == 'Hoja2':
        ws['A1'].value = "ESTADO DE RESULTADOS"
        FILAS_GRISES = [8,10,16,28,32]

    elif ws.title == 'Hoja3':
        ws['A1'].value = "ESTADO DE PATRIMONIO NETO"
        ws.column_dimensions[get_column_letter(1)].width = 20
        for col in range(2, ws.max_column+1):
            ws.column_dimensions[get_column_letter(col)].width = 45
        FILAS_GRISES = [8,11,16,30,31,32,35,40,54,55,56,59,64,78,79,80,83,88,102,103,104,107,112,126,127,128,131,136,150,151]

    elif ws.title == 'Hoja4':
        ws['A1'].value = "ESTADO DE FLUJO DE EFECTIVO"
        FILAS_GRISES = [8,9,10,18,28,47,48,49,62,75,76,77,84,95]
    return FILAS_GRISES

def FormatoSituacionFinanciera(nroHoja, nombre):
    ws = nroHoja
    nombreEmpresa = nombre

    max_row = ws.max_row
    max_col = ws.max_column
    
    # Ajustar el ancho de las columnas
    ws.column_dimensions[get_column_letter(1)].width = 2
    ws.column_dimensions[get_column_letter(2)].width = 2
    ws.column_dimensions[get_column_letter(3)].width = 45
    for i in range(4, max_col + 5):
        ws.column_dimensions[get_column_letter(i)].width = 14.5

    # Título Principal (Fila 1, Columna A)
    ws['A1'].font = fuente_titulo
    ws['A3'].value = "Periodo: Anual"
    ws['A3'].font = negrita
    ws['A4'].value = f"Empresa: {nombreEmpresa}"
    ws['A4'].font = negrita
    ws['A5'].value = "Tipo: Individual"
    ws['A5'].font = negrita

    # Ajustamos la fila de los encabezados (ej. Fila 6 del Excel)
    DATA_START_ROW = 7
    FILAS_GRISES = hojas(ws)

    HEADER_ROW = 7
    for col in range(3, max_col + 5):
        cell = ws.cell(row=HEADER_ROW, column=col) 
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for row in range(DATA_START_ROW, max_row + 1):
        for col in range(3, max_col + 5):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.font = Contenido

            if col == 3:
                cell.alignment = Alignment(horizontal='left', vertical='center')

            if row in FILAS_GRISES:
                cell.alignment = Alignment(horizontal='left', vertical='center')
                cell.fill = SECTION_FILL

                if col in range(4, max_col + 5):
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if col in range(4, max_col + 5):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if row == 9:
                cell = ws.cell(row=DATA_START_ROW, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center')

def FormatoResultados(nroHoja, nombre):
    ws = nroHoja
    nombreEmpresa = nombre

    max_row = ws.max_row
    max_col = ws.max_column
    
    # Ajustar el ancho de las columnas
    ws.column_dimensions[get_column_letter(1)].width = 2
    ws.column_dimensions[get_column_letter(2)].width = 2
    ws.column_dimensions[get_column_letter(3)].width = 45
    for i in range(4, max_col + 5):
        ws.column_dimensions[get_column_letter(i)].width = 14.5

    # Título Principal (Fila 1, Columna A)
    ws['A1'].font = fuente_titulo
    ws['A3'].value = "Periodo: Anual"
    ws['A3'].font = negrita
    ws['A4'].value = f"Empresa: {nombreEmpresa}"
    ws['A4'].font = negrita
    ws['A5'].value = "Tipo: Individual"
    ws['A5'].font = negrita

    # Ajustamos la fila de los encabezados (ej. Fila 6 del Excel)
    DATA_START_ROW = 7
    
    FILAS_GRISES = hojas(ws)

    for row in range(DATA_START_ROW, max_row - 14):
        for col in range(3, max_col + 5):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.font = Contenido

            if col == 3:
                cell.alignment = Alignment(horizontal='left', vertical='center')

            if row in FILAS_GRISES:
                cell.alignment = Alignment(horizontal='left', vertical='center')
                cell.fill = SECTION_FILL

                if col in range(4, max_col + 5):
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if col in range(4, max_col + 5):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if row == 9:
                cell = ws.cell(row=DATA_START_ROW, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center')

    limpiar_rango_Formato(ws, 'C33:I47')

def FormatoFlujoEfectivo(nroHoja, nombre):
    ws = nroHoja
    nombreEmpresa = nombre

    max_row = ws.max_row
    max_col = ws.max_column
    
    # Ajustar el ancho de las columnas
    ws.column_dimensions[get_column_letter(1)].width = 2
    ws.column_dimensions[get_column_letter(2)].width = 2
    ws.column_dimensions[get_column_letter(3)].width = 45
    for i in range(4, max_col + 5):
        ws.column_dimensions[get_column_letter(i)].width = 14.5

    # Título Principal (Fila 1, Columna A)
    ws['A1'].font = fuente_titulo
    ws['A3'].value = "Periodo: Anual"
    ws['A3'].font = negrita
    ws['A4'].value = f"Empresa: {nombreEmpresa}"
    ws['A4'].font = negrita
    ws['A5'].value = "Tipo: Individual"
    ws['A5'].font = negrita

    # Ajustamos la fila de los encabezados (ej. Fila 6 del Excel)
    DATA_START_ROW = 7
    
    FILAS_GRISES = hojas(ws)

    for row in range(DATA_START_ROW, max_row - 4):
        for col in range(3, max_col + 5):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.font = Contenido

            if col == 3:
                cell.alignment = Alignment(horizontal='left', vertical='center')

            if row in FILAS_GRISES:
                cell.alignment = Alignment(horizontal='left', vertical='center')
                cell.fill = SECTION_FILL

                if col in range(4, max_col + 5):
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if col in range(4, max_col + 5):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if row == 9:
                cell = ws.cell(row=DATA_START_ROW, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center')
    limpiar_rango_Formato(ws, "C96:I100")

def FormatoPatrimonio(wb, nroHoja, nombre):
    ws = nroHoja
    nombreEmpresa = nombre

    max_row = ws.max_row
    max_col = ws.max_column

    ReordenarTabla(ws, max_col)
    
    # Ajustar el ancho de las columnas
    ws.column_dimensions[get_column_letter(4)].width = 14.5
    for i in range(5, max_col + 5):
        ws.column_dimensions[get_column_letter(i)].width = 10

    # Título Principal (Fila 1, Columna A)
    ws['A1'].font = fuente_titulo
    ws['A3'].value = "Periodo: Anual"
    ws['A3'].font = negrita
    ws['A4'].value = f"Empresa: {nombreEmpresa}"
    ws['A4'].font = negrita
    ws['A5'].value = "Tipo: Individual"
    ws['A5'].font = negrita

    # Ajustamos la fila de los encabezados (ej. Fila 6 del Excel)
    DATA_START_ROW = 7
    
    FILAS_GRISES = hojas(ws)

    for row in range(DATA_START_ROW, max_row + 97):
        for col in range(3, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER
            cell.font = Contenido

            if col == 3:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            if row in FILAS_GRISES:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL

                if col in range(4, max_col + 1):
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if col in range(4, max_col + 1):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.number_format = FORMATO_NUMERICO_FINANCIERO
            
            if row == 9:
                cell = ws.cell(row=DATA_START_ROW, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    ws.column_dimensions[get_column_letter(1)].width = 5
    ws.column_dimensions[get_column_letter(2)].width = 5
    ws.column_dimensions[get_column_letter(3)].width = 10

    ws.merge_cells('B8:B31')    # COMBINAR Y CENTRAR - AÑO 2024
    celda2024 = ws['B8']
    celda2024.value = "2024"
    encabezadosFechasVerticales(celda2024)

    ws.merge_cells('B32:B55')    # COMBINAR Y CENTRAR - AÑO 2023
    celda2023 = ws['B32']
    celda2023.value = "2023"
    encabezadosFechasVerticales(celda2023)

    ws.merge_cells('B56:B79')    # COMBINAR Y CENTRAR - AÑO 2022
    celda2022 = ws['B56']
    celda2022.value = "2022"
    encabezadosFechasVerticales(celda2022)

    ws.merge_cells('B80:B103')    # COMBINAR Y CENTRAR - AÑO 2021
    celda2021 = ws['B80']
    celda2021.value = "2021"
    encabezadosFechasVerticales(celda2021)

    ws.merge_cells('B104:B127')    # COMBINAR Y CENTRAR - AÑO 2020
    celda2020 = ws['B104']
    celda2020.value = "2020"
    encabezadosFechasVerticales(celda2020)

def limpiar_rango_Formato(ws, rango_excel):
    min_col, min_row, max_col, max_row = range_boundaries(rango_excel)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.value = None

def encabezadosFechasVerticales(celda):
    celda.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    celda.font = Font(size=16, bold=True)
    celda.alignment = Alignment(
        horizontal='center',
        vertical='center',
        text_rotation=90  # Rota el texto 90 grados (hacia arriba)
    )
    celda.border = THIN_BORDER

def ReordenarTabla(ws, max_col):    
    # 1. Definir los límites
    # Usaremos A y la última columna (ej. E)
    MAX_COL_LETTER = get_column_letter(max_col + 1)
    MAX_ROW = ws.max_row

    RANGO_SUPERIOR = f"C8:{MAX_COL_LETTER}31"
    ws.move_range(
        RANGO_SUPERIOR,
        rows=+48,
        cols=0
    )

    RANGO_INFERIOR = f"C32:{MAX_COL_LETTER}{MAX_ROW}"
    ws.move_range(
        RANGO_INFERIOR, 
        rows=-24,
        cols=0
    )

    RANGO_SUPERIOR2 = f"C56:{MAX_COL_LETTER}79"
    ws.move_range(
        RANGO_SUPERIOR2,
        rows=-24,
        cols=0
    )

def union_archivos(path_xlsx_origen, path_xlsx_destino, columna):
    # 1. CARGAR AMBOS WORKBOOKS
    wb_origen = load_workbook(path_xlsx_origen, data_only=True)
    wb_destino = load_workbook(path_xlsx_destino)
    
    # 2. DEFINIR HOJAS DE ORIGEN Y DESTINO
    ws_Origen_hoja1 = wb_origen['Hoja1']
    ws_Destino_hoja1 = wb_destino['Hoja1']
    ws_Origen_hoja2 = wb_origen['Hoja2']
    ws_Destino_hoja2 = wb_destino['Hoja2']
    ws_Origen_hoja3 = wb_origen['Hoja3']
    ws_Destino_hoja3 = wb_destino['Hoja3']
    ws_Origen_hoja4 = wb_origen['Hoja4']
    ws_Destino_hoja4 = wb_destino['Hoja4']
    
    # 3. DEFINIR RANGO Y POSICIÓN
    rango_a_copiar_2 = 'C8:AA55'
    fila_destino_inicial = 7
    fila_destino_inicial_2 = 56
    fila_destino_inicial_3 = 104
    columna_destino_inicial = columna

    # 4. EJECUTAR COPIA
    #UNION DE SITUACION FINANCIERA - HOJA 1
    copiar_celdas(
        ws_Origen_hoja1,
        ws_Destino_hoja1,
        'D7:E83',
        fila_destino_inicial,
        columna_destino_inicial
    )

    #UNION DE RESULTADOS - HOJA 2
    copiar_celdas(
        ws_Origen_hoja2,
        ws_Destino_hoja2,
        'D7:E83',
        fila_destino_inicial,
        columna_destino_inicial
    )

    #UNION DE PATRIMONIO - HOJA 3
    """
    if columna == 5:
        copiar_celdas(
            ws_Origen_hoja3,
            ws_Destino_hoja3,
            rango_a_copiar_2,
            fila_destino_inicial_2,
            columna_destino_inicial - 3
        )
    else:
    
        copiar_celdas(
            ws_Origen_hoja3,
            ws_Destino_hoja3,
            rango_a_copiar_2,
            fila_destino_inicial_3,
            columna_destino_inicial - 5
        )
    """

    #UNION DE FLUJO DE EFECTIVOS - HOJA 4
    copiar_celdas(
        ws_Origen_hoja4,
        ws_Destino_hoja4,
        'D7:E95',
        fila_destino_inicial,
        columna_destino_inicial
    )
    
    # 5. GUARDAR EL ARCHIVO DESTINO (EL ARCHIVO ORIGEN NO SE MODIFICA)
    wb_destino.save(path_xlsx_destino)

def copiar_celdas(ws_origen, ws_destino, rango_origen: str, fila_inicio_destino: int, columna_inicio_destino: int):    
    # Obtener las coordenadas del rango de origen
    try:
        min_col, min_row, max_col, max_row = range_boundaries(rango_origen) 
    except ValueError:
        print(f"Error: Rango '{rango_origen}' no es válido.")
        return

    fila_destino = fila_inicio_destino
    
    # Iterar sobre las filas y columnas del rango de origen
    for row in ws_origen.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        col_destino = columna_inicio_destino
        
        for cell_origen in row:
            # Obtener la celda de destino
            cell_destino = ws_destino.cell(row=fila_destino, column=col_destino)
            
            # 1. Copiar Valor: Solo si la celda tiene un valor (ignora celdas combinadas secundarias)
            if cell_origen.value is not None:
                cell_destino.value = cell_origen.value 
                
                # 2. Copiar Formato Numérico
                if cell_origen.number_format:
                    cell_destino.number_format = FORMATO_NUMERICO_FINANCIERO
            else:
                # Opcional: Asegurar que la celda destino también esté vacía si el origen lo está
                cell_destino.value = None

            col_destino += 1
        fila_destino += 1

def analisis_VH(path_xlsx):
    wb = load_workbook(path_xlsx)

    try:
        Ratios = wb['Hoja5']
    except KeyError:
        Ratios = wb.copy_worksheet(wb['Hoja1'])
        Ratios.title = "Hoja5"

    Ratios['A1'].value = "RATIOS FINANCIEROS"

    sistFinan = wb['Hoja1']
    resultados = wb['Hoja2']
    flujos = wb['Hoja4']
    FormatoAnalisis1(sistFinan)
    FormatoAnalisis2(resultados)
    FormatoAnalisis3(flujos)

    wb.save(path_xlsx)

def FormatoAnalisis1(ws):
    ws.column_dimensions[get_column_letter(9)].width = 3
    ws.column_dimensions[get_column_letter(15)].width = 3
    ws.merge_cells('J6:N6')
    aplicarBorde(ws, 'J6:N83')
    ws['J6'].value = "Análisis Vertical"
    ws['J6'].fill = ENCABEZADO_NARANJA
    ws['J6'].font = negrita
    ws['J6'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('P6:S6')
    aplicarBorde(ws, 'P6:S83')
    ws['P6'].value = "Análisis Horizontal"
    ws['P6'].fill = ENCABEZADO_NARANJA
    ws['P6'].font = negrita
    ws['P6'].alignment = Alignment(horizontal='center', vertical='center')

    copiar_celdas(ws,ws,'D7:H7',7,10)
    copiar_celdas(ws,ws,'D7:G7',7,16)

    for row in range(7, 84):
        for col in range(10, 15):
            cell = ws.cell(row=row, column=col)
            if row == 7:
                cell = ws.cell(row=7, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            elif row in [8,9,24,25,43,44,45,46,59,60,72,73,74,82,83]:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL
            
            if row in range(8, 84):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for col in range(16, 20):
            cell = ws.cell(row=row, column=col)
            if row == 7:
                cell = ws.cell(row=7, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            elif row in [8,9,24,25,43,44,45,46,59,60,72,73,74,82,83]:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL
            
            if row in range(8, 84):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def FormatoAnalisis2(ws):
    ws.column_dimensions[get_column_letter(9)].width = 3
    ws.column_dimensions[get_column_letter(15)].width = 3
    ws.merge_cells('J6:N6')
    aplicarBorde(ws, 'J6:N32')
    ws['J6'].value = "Análisis Vertical"
    ws['J6'].fill = ENCABEZADO_NARANJA
    ws['J6'].font = negrita
    ws['J6'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('P6:S6')
    aplicarBorde(ws, 'P6:S32')
    ws['P6'].value = "Análisis Horizontal"
    ws['P6'].fill = ENCABEZADO_NARANJA
    ws['P6'].font = negrita
    ws['P6'].alignment = Alignment(horizontal='center', vertical='center')

    copiar_celdas(ws,ws,'D7:H7',7,10)
    copiar_celdas(ws,ws,'D7:G7',7,16)

    for row in range(7, 33):
        for col in range(10, 15):
            cell = ws.cell(row=row, column=col)
            if row == 7:
                cell = ws.cell(row=7, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            elif row in [8,10,16,28,32]:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL
            
            if row in range(8, 33):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        for col in range(16, 20):
            cell = ws.cell(row=row, column=col)
            if row == 7:
                cell = ws.cell(row=7, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            elif row in [8,10,16,28,32]:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL
            
            if row in range(8, 33):
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def FormatoAnalisis3(ws):
    ws.column_dimensions[get_column_letter(9)].width = 3
    ws.column_dimensions[get_column_letter(15)].width = 3
    ws.merge_cells('J6:N6')
    aplicarBorde(ws, 'J6:N95')
    ws['J6'].value = "Análisis Vertical"
    ws['J6'].fill = ENCABEZADO_NARANJA
    ws['J6'].font = negrita
    ws['J6'].alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('P6:S6')
    aplicarBorde(ws, 'P6:S95')
    ws['P6'].value = "Análisis Horizontal"
    ws['P6'].fill = ENCABEZADO_NARANJA
    ws['P6'].font = negrita
    ws['P6'].alignment = Alignment(horizontal='center', vertical='center')

    copiar_celdas(ws,ws,'D7:H7',7,10)
    copiar_celdas(ws,ws,'D7:G7',7,16)

    for row in range(7, 96):
        for col in range(10, 15):
            cell = ws.cell(row=row, column=col)
            if row == 7:
                cell = ws.cell(row=7, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            elif row in [8,9,10,18,28,47,48,49,62,75,76,77,84,95]:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL
        
        for col in range(16, 20):
            cell = ws.cell(row=row, column=col)
            if row == 7:
                cell = ws.cell(row=7, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            elif row in [8,9,10,18,28,47,48,49,62,75,76,77,84,95]:
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.fill = SECTION_FILL

def limpiar_rango_Libre(ws, rango_excel):
    BORDE_POR_DEFECTO = Border(left=Side(style=None), 
                               right=Side(style=None), 
                               top=Side(style=None), 
                               bottom=Side(style=None))
    try:
        min_col, min_row, max_col, max_row = range_boundaries(rango_excel)
    except Exception:
        print(f"Error: El rango '{rango_excel}' no es un formato de rango válido.")
        return

    # 3. Iterar y limpiar cada celda
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = BORDE_POR_DEFECTO
            cell.number_format = 'General'
            
def aplicarBorde(ws, rango_excel):
    min_col, min_row, max_col, max_row = range_boundaries(rango_excel)

    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = THIN_BORDER

def analisis_Ratios(path_xlsx):
    wb = load_workbook(path_xlsx)
    ws = wb['Hoja5']

    ws.column_dimensions[get_column_letter(9)].width = 3
    ws.column_dimensions[get_column_letter(10)].width = 30
    copiar_celdas(ws,ws,'D7:H7',7,11)
    copiar_celdas(ws,ws,'D7:H7',11,11)
    copiar_celdas(ws,ws,'D7:H7',16,11)
    copiar_celdas(ws,ws,'D7:H7',20,11)

    ws['J7'].value = "RATIOS DE LIQUIDEZ"
    ws['J7'].fill = ENCABEZADO_NARANJA
    ws['J7'].font = Font(size=11, bold=True)
    ws['J8'].value = "Liquidez Corriente"
    ws['J9'].value = "Prueba Ácida"

    ws['J11'].value = "RATIOS DE GESTIÓN"
    ws['J11'].fill = ENCABEZADO_NARANJA
    ws['J11'].font = Font(size=11, bold=True)
    ws['J12'].value = "Rotación de Cuentas por cobrar"
    ws['J13'].value = "Rotación de Inventarios"
    ws['J14'].value = "Rotación de Activos Totales"

    ws['J16'].value = "RATIOS DE ENDEUDAMIENTO"
    ws['J16'].fill = ENCABEZADO_NARANJA
    ws['J16'].font = Font(size=11, bold=True)
    ws['J17'].value = "Razón de deuda total"
    ws['J18'].value = "Razón de deuda/patrimonio"

    ws['J20'].value = "RATIOS DE RENTABILIDAD"
    ws['J20'].fill = ENCABEZADO_NARANJA
    ws['J20'].font = Font(size=11, bold=True)
    ws['J21'].value = "Margen neto"
    ws['J22'].value = "ROA"
    ws['J23'].value = "ROE"

    aplicarBorde(ws, 'J7:O9')
    aplicarBorde(ws, 'J11:O14')
    aplicarBorde(ws, 'J16:O18')
    aplicarBorde(ws, 'J20:O23')

    for row in range(6, 24):
        for col in range(11, 16):
            cell = ws.cell(row=row, column=col)
            if row in [7,11,16,20]:
                cell = ws.cell(row=row, column=col) 
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.number_format = FORMATO_NUMERICO
            elif row in [8,9,11,12,13]:
                cell.number_format = FORMATO_NUMERICO_FINANCIERO
            else:
                cell.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES
    wb.save(path_xlsx)

def graficosRatios(path_xlsx):
    wb = load_workbook(path_xlsx)
    try:
        GraRati = wb['Hoja6']
    except KeyError:
        GraRati = wb.create_sheet(title="Hoja6", index=None)
    ws = GraRati

    dir_path = os.path.dirname(path_xlsx)
    nombre = os.path.basename(dir_path)
    nombreEmpresa = nombre.replace('_', ' ')

    ws['A1'].value = "GRÁFICOS DE RATIOS"
    ws['A1'].font = fuente_titulo
    ws['A3'].value = "Periodo: Anual"
    ws['A3'].font = negrita
    ws['A4'].value = f"Empresa: {nombreEmpresa}"
    ws['A4'].font = negrita
    ws['A5'].value = "Tipo: Individual"
    ws['A5'].font = negrita

    ws['C7'].value = "RATIOS DE LIQUIDEZ"

    ws.column_dimensions[get_column_letter(1)].width = 3
    ws.column_dimensions[get_column_letter(2)].width = 3
    for i in range(3,30):
        ws.column_dimensions[get_column_letter(i)].width = 15
    for i in [5,8,11,14]:
        ws.column_dimensions[get_column_letter(i)].width = 32

    ws.merge_cells('C7:G7')
    ws.merge_cells('E8:E14')
    ws['C7'].value = "RATIOS DE LIQUIDEZ"
    ws['C7'].fill = ENCABEZADO_PURPURA
    ws['C7'].font = HEADER_FONT
    ws['C7'].font = Font(size=12, bold=True)
    ws.merge_cells('C8:D8')
    ws['C8'].value = "Liquidez Corriente"
    ws['C8'].fill = ENCABEZADO_NARANJA
    ws['C9'].value = "Año"
    ws['C9'].fill =ENCABEZADO_CELESTE
    ws['D9'].value = "Valor"
    ws['D9'].fill =ENCABEZADO_CELESTE
    ws.merge_cells('F8:G8')
    ws['F8'].value = "Prueba Ácida"
    ws['F8'].fill = ENCABEZADO_NARANJA
    ws['F9'].value = "Año"
    ws['F9'].fill =ENCABEZADO_CELESTE
    ws['G9'].value = "Valor"
    ws['G9'].fill =ENCABEZADO_CELESTE

    ws.merge_cells('I7:P7')
    ws.merge_cells('K8:K13')
    ws.merge_cells('N8:N13')
    ws['I7'].value = "RATIOS DE GESTIÓN"
    ws['I7'].fill = ENCABEZADO_PURPURA
    ws['I7'].font = HEADER_FONT
    ws['I7'].font = Font(size=12, bold=True)
    ws.merge_cells('I8:J8')
    ws['I8'].value = "Rotación de Cuentas por cobrar"
    ws['I8'].fill = ENCABEZADO_NARANJA
    ws['I9'].value = "Año"
    ws['I9'].fill =ENCABEZADO_CELESTE
    ws['J9'].value = "Valor"
    ws['J9'].fill =ENCABEZADO_CELESTE
    ws.merge_cells('L8:M8')
    ws['L8'].value = "Rotación de Inventarios"
    ws['L8'].fill = ENCABEZADO_NARANJA
    ws['L9'].value = "Año"
    ws['L9'].fill =ENCABEZADO_CELESTE
    ws['M9'].value = "Valor"
    ws['M9'].fill =ENCABEZADO_CELESTE
    ws.merge_cells('O8:P8')
    ws['O8'].value = "Rotación de Activos Totales"
    ws['O8'].fill = ENCABEZADO_NARANJA
    ws['O9'].value = "Año"
    ws['O9'].fill = ENCABEZADO_CELESTE
    ws['P9'].value = "Valor"
    ws['P9'].fill = ENCABEZADO_CELESTE

    ws.merge_cells('C30:G30')
    ws.merge_cells('E31:E37')
    ws['C30'].value = "RATIOS DE ENDEUDAMIENTO"
    ws['C30'].fill = ENCABEZADO_PURPURA
    ws['C30'].font = HEADER_FONT
    ws['C30'].font = Font(size=12, bold=True)
    ws.merge_cells('C31:D31')
    ws['C31'].value = "Razón de deuda total"
    ws['C31'].fill = ENCABEZADO_NARANJA
    ws['C32'].value = "Año"
    ws['C32'].fill = ENCABEZADO_CELESTE
    ws['D32'].value = "Valor"
    ws['D32'].fill = ENCABEZADO_CELESTE
    ws.merge_cells('F31:G31')
    ws['F31'].value = "Razón de deuda/patrimonio"
    ws['F31'].fill = ENCABEZADO_NARANJA
    ws['F32'].value = "Año"
    ws['F32'].fill = ENCABEZADO_CELESTE
    ws['G32'].value = "Valor"
    ws['G32'].fill = ENCABEZADO_CELESTE

    ws.merge_cells('I30:P30')
    ws.merge_cells('K31:K37')
    ws.merge_cells('N31:N37')
    ws['I30'].value = "RATIOS DE RENTABILIDAD"
    ws['I30'].fill = ENCABEZADO_PURPURA
    ws['I30'].font = HEADER_FONT
    ws['I30'].font = Font(size=12, bold=True)
    ws.merge_cells('I31:J31')
    ws['I31'].value = "Margen neto"
    ws['I31'].fill = ENCABEZADO_NARANJA
    ws['I32'].value = "Año"
    ws['I32'].fill = ENCABEZADO_CELESTE
    ws['J32'].value = "Valor"
    ws['J32'].fill = ENCABEZADO_CELESTE
    ws.merge_cells('L31:M31')
    ws['L31'].value = "ROA"
    ws['L31'].fill = ENCABEZADO_NARANJA
    ws['L32'].value = "Año"
    ws['L32'].fill = ENCABEZADO_CELESTE
    ws['M32'].value = "Valor"
    ws['M32'].fill = ENCABEZADO_CELESTE
    ws.merge_cells('O31:P31')
    ws['O31'].value = "ROE"
    ws['O31'].fill = ENCABEZADO_NARANJA
    ws['O32'].value = "Año"
    ws['O32'].fill = ENCABEZADO_CELESTE
    ws['P32'].value = "Valor"
    ws['P32'].fill = ENCABEZADO_CELESTE

    aplicarBorde(ws, 'C7:G14')
    centrar_rango(ws, 'C7:G14')
    aplicarBorde(ws, 'I7:P13')
    centrar_rango(ws, 'I7:P13')
    aplicarBorde(ws, 'C30:G37')
    centrar_rango(ws, 'C30:G37')
    aplicarBorde(ws, 'I30:P37')
    centrar_rango(ws, 'I30:P37')

    wb.save(path_xlsx)
    
def renombrar(path_xlsx):
    wb = load_workbook(path_xlsx)
    sistFinan = wb['Hoja1']
    resultados = wb['Hoja2']
    patrimonio = wb['Hoja3']
    flujos = wb['Hoja4']
    ratios = wb['Hoja5']
    graratios = wb['Hoja6']

    sistFinan.title = "Estado de Situación Financiera"
    resultados.title = "Estado de Resultados"
    patrimonio.title = "Estado de Patrimonio Neto"
    flujos.title = "Estado de Flujo de Efectivo"
    ratios.title = "Ratios Financieros"
    graratios.title = "Gráficos de Ratios Financieros"
    wb.save(path_xlsx)

def centrar_rango(ws, rango):
    alineacion_centrada = Alignment(horizontal='center', vertical='center')
    min_col, min_row, max_col, max_row = range_boundaries(rango)
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.alignment = alineacion_centrada

def valor(wb, nombre_hoja, celda):
    valor = wb[nombre_hoja][celda].value
    if valor is None:
        return "0"
    s = str(valor).strip()
    # Manejar formato contable con paréntesis (negativos)
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]

    # Normalizar separadores para convertir a número
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        num = float(s)
        # Formatear salida con punto miles y coma decimal
        if num.is_integer():
            return f"{int(num):,}".replace(",", ".")  # enteros sin decimales
        else:
            return f"{num:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except ValueError:
        return "0"

def analisisVertical(path_xlsx):
    wb = load_workbook(path_xlsx)
    ws1 = wb['Hoja1']
    ws2 = wb['Hoja2']
    ws4 = wb['Hoja4']

    for row in range(8, 84):    # SITUACIÓN FINANCIERA
        for offset, col in enumerate(range(4, 9)):
            num = ws1.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws1.cell(row=44, column=col).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = num / den
            else:
                resultado = 0.00

            celda_resultado = ws1.cell(row=row, column=10 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for row in range(8, 33):    # RESULTADOS
        for offset, col in enumerate(range(4, 9)):
            num = ws2.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws2.cell(row=8, column=col).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = num / den
            else:
                resultado = 0.00

            celda_resultado = ws2.cell(row=row, column=10 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for row in range(8, 48):    # FLUJO DE EFECTIVO
        for offset, col in enumerate(range(4, 9)):
            num = ws4.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws4.cell(row=47, column=col).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = num / den
            else:
                resultado = 0.00

            celda_resultado = ws4.cell(row=row, column=10 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES

    for row in range(48, 76):    # FLUJO DE EFECTIVO
        for offset, col in enumerate(range(4, 9)):
            num = ws4.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws4.cell(row=75, column=col).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = num / den
            else:
                resultado = 0.00

            celda_resultado = ws4.cell(row=row, column=10 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for row in range(77, 96):    # FLUJO DE EFECTIVO
        for offset, col in enumerate(range(4, 9)):
            num = ws4.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws4.cell(row=95, column=col).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = num / den
            else:
                resultado = 0.00

            celda_resultado = ws4.cell(row=row, column=10 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES

    wb.save(path_xlsx)

def analisisHorizontal(path_xlsx):
    wb = load_workbook(path_xlsx)
    ws1 = wb['Hoja1']
    ws2 = wb['Hoja2']
    ws4 = wb['Hoja4']

    for row in range(8, 84):    # SITUACIÓN FINANCIERA
        for offset, col in enumerate(range(4, 8)):
            num = ws1.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws1.cell(row=row, column=col+1).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = (num - den) / den
            else:
                resultado = 0.00

            celda_resultado = ws1.cell(row=row, column=16 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES

    for row in range(8, 33):    # RESULTADOS
        for offset, col in enumerate(range(4, 8)):
            num = ws2.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws2.cell(row=row, column=col+1).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = (num - den) / den
            else:
                resultado = 0.00

            celda_resultado = ws2.cell(row=row, column=16 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES

    for row in range(8, 96):    # FLUJO DE EFECTIVO
        for offset, col in enumerate(range(4, 8)):
            num = ws4.cell(row=row, column=col).value
            num = convertir_a_numero(num)
            den = ws4.cell(row=row, column=col+1).value
            den = convertir_a_numero(den)

            if den not in (0, None):
                resultado = (num - den) / den
            else:
                resultado = 0.00

            celda_resultado = ws4.cell(row=row, column=16 + offset, value=resultado)
            celda_resultado.value = resultado
            celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES

    wb.save(path_xlsx)

def analisisRatiosCalculo(path_xlsx):
    wb = load_workbook(path_xlsx)
    ws1 = wb['Hoja5']
    ws2 = wb['Hoja2']

    for offset, col in enumerate(range(4, 9)):  # LIQUIDEZ CORRIENTE
        num = ws1.cell(row=24, column=col).value
        num = convertir_a_numero(num)
        den = ws1.cell(row=59, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=8, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_NUMERICO_FINANCIERO


    for offset, col in enumerate(range(4, 9)):  # PRUEBA ÁCIDA
        num1 = ws1.cell(row=24, column=col).value
        num1 = convertir_a_numero(num1)
        num2 = ws1.cell(row=17, column=col).value
        num2 = convertir_a_numero(num2)
        den = ws1.cell(row=59, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = (num1 - num2) / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=9, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_NUMERICO_FINANCIERO

    
    for offset, col in enumerate(range(4, 9)):  # RAZÓN DE DEUDA TOTAL
        num = ws1.cell(row=73, column=col).value
        num = convertir_a_numero(num)
        den = ws1.cell(row=44, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=17, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for offset, col in enumerate(range(4, 9)):  # RAZÓN DE DEUDA/PATRIMONIO
        num = ws1.cell(row=73, column=col).value
        num = convertir_a_numero(num)
        den = ws1.cell(row=82, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=18, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for offset, col in enumerate(range(4, 9)):  # MARGEN NETO
        num = ws2.cell(row=32, column=col).value
        num = convertir_a_numero(num)
        den = ws2.cell(row=8, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=21, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for offset, col in enumerate(range(4, 9)):  # ROA
        num = ws2.cell(row=32, column=col).value
        num = convertir_a_numero(num)
        den = ws1.cell(row=44, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=22, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for offset, col in enumerate(range(4, 9)):  # ROE
        num = ws2.cell(row=32, column=col).value
        num = convertir_a_numero(num)
        den = ws1.cell(row=82, column=col).value
        den = convertir_a_numero(den)

        if den not in (0, None):
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=23, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_PORCENTAJE_DOS_DECIMALES


    for offset, col in enumerate(range(4, 8)):  # ROTACIÓN DE CUENTAS POR COBRAR
        num = ws2.cell(row=8, column=col).value
        num = convertir_a_numero(num)

        # Valores columna col (D-G)
        d1 = convertir_a_numero(ws1.cell(row=12, column=col).value)
        d2 = convertir_a_numero(ws1.cell(row=13, column=col).value)
        d3 = convertir_a_numero(ws1.cell(row=14, column=col).value)
        d4 = convertir_a_numero(ws1.cell(row=15, column=col).value)
        d5 = convertir_a_numero(ws1.cell(row=28, column=col).value)
        d6 = convertir_a_numero(ws1.cell(row=29, column=col).value)
        d7 = convertir_a_numero(ws1.cell(row=30, column=col).value)
        d8 = convertir_a_numero(ws1.cell(row=31, column=col).value)

        # Valores columna col+1 (E-H)
        d9  = convertir_a_numero(ws1.cell(row=12, column=col+1).value)
        d10 = convertir_a_numero(ws1.cell(row=13, column=col+1).value)
        d11 = convertir_a_numero(ws1.cell(row=14, column=col+1).value)
        d12 = convertir_a_numero(ws1.cell(row=15, column=col+1).value)
        d13 = convertir_a_numero(ws1.cell(row=28, column=col+1).value)
        d14 = convertir_a_numero(ws1.cell(row=29, column=col+1).value)
        d15 = convertir_a_numero(ws1.cell(row=30, column=col+1).value)
        d16 = convertir_a_numero(ws1.cell(row=31, column=col+1).value)

        den = (d1+d2+d3+d4+d5+d6+d7+d8 + d9+d10+d11+d12+d13+d14+d15+d16) / 2

        if den != 0:
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=12, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_NUMERICO_FINANCIERO
    ws1['O12'].value = 0.00


    for offset, col in enumerate(range(4, 8)):  # ROTACIÓN DE INVENTARIOS
        num = ws2.cell(row=9, column=col).value
        num = convertir_a_numero(num)

        d1 = convertir_a_numero(ws1.cell(row=17, column=col).value)
        d2 = convertir_a_numero(ws1.cell(row=33, column=col+1).value)
        d3 = convertir_a_numero(ws1.cell(row=17, column=col).value)
        d4 = convertir_a_numero(ws1.cell(row=33, column=col+1).value)

        den = (d1+d2+d3+d4) / 2

        if den != 0:
            resultado = (num / den) * -1
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=13, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_NUMERICO_FINANCIERO
    ws1['O13'].value = 0.00


    for offset, col in enumerate(range(4, 9)):  # ROTACIÓN DE ACTIVOS TOTALES
        num = ws2.cell(row=8, column=col).value
        num = convertir_a_numero(num)

        den = convertir_a_numero(ws1.cell(row=44, column=col).value)

        if den != 0:
            resultado = num / den
        else:
            resultado = 0.00

        celda_resultado = ws1.cell(row=14, column=11 + offset, value=resultado)
        celda_resultado.value = resultado
        celda_resultado.number_format = FORMATO_NUMERICO_FINANCIERO


    

    wb.save(path_xlsx)

def valor(wb, nombre_hoja, celda):
    valor = wb[nombre_hoja][celda].value
    return convertir_a_numero(valor)

def convertir_a_numero(valor): 
    if valor is None: 
        return 0.0 
    s = str(valor).strip() 

    # Manejar paréntesis (negativos contables) 

    if s.startswith("(") and s.endswith(")"): 
        s = "-" + s[1:-1].strip() 
    
    # Eliminar separadores de miles (,) y mantener el punto como decimal
    s = s.replace(",", "") 
    try: 
        return float(s) 
    except ValueError: 
        return 0.0