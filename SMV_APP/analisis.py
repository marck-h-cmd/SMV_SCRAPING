import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter

#gogo
# --- Definición de Estilos ---
# Color de relleno de encabezado (Azul Oscuro/Teal)
HEADER_FILL = PatternFill(start_color="003366", end_color="003366", fill_type="solid")

# Color de relleno para la sección principal (Gris Claro)
SECTION_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

# Fuente para encabezados (Blanco y Negrita)
HEADER_FONT = Font(bold=True, color="FFFFFF")

# Fuente para el título "ESTADO DE SITUACION FINANCIERA"
TITLE_FONT = Font(bold=True)

# Estilo de borde (borde fino para todas las celdas de datos)
THIN_SIDE = Side(style='thin', color="000000")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
# -----------------------------

CARPETA_FINANCIEROS = r"C:/Users/FELIX/Desktop/CICLO 6/FINANZAS CORPORATIVAS/PROYECTO/SMV_SCRAPING/descargas_smv/ADMINISTRADORA_JOCKEY_PLAZA_SHOPPING_CENTER_SA"

def marcar_celda_roja(celda_marcada="B2"):
    resultados = []

    for archivo in os.listdir(CARPETA_FINANCIEROS):
        if archivo.endswith(".xls"):
            path_xls = os.path.join(CARPETA_FINANCIEROS, archivo)
            path_xlsx = path_xls + "x"

            tablas = pd.read_html(path_xls)
            if not tablas:
                continue

            # 1. Guardar DataFrames en XLSX (sin estilos)
            # Usamos `header=False` para tener más control sobre la Fila 1 y los encabezados
            with pd.ExcelWriter(path_xlsx, engine="openpyxl") as writer:
                for i, df in enumerate(tablas, start=1):
                    # Eliminamos las primeras filas que contienen los datos de la empresa/fecha
                    df_cleaned = df.iloc[5:] # Asumiendo que las primeras 5 filas son metadata
                    df_cleaned.to_excel(writer, sheet_name=f"Hoja{i}", index=False, header=False, startrow=2)

            # 2. Abrir con openpyxl para aplicar los estilos manuales
            wb = load_workbook(path_xlsx)
            ws = wb.active # Primera hoja (Hoja1)

            # Determinar el rango de la tabla de datos
            max_row = ws.max_row
            max_col = ws.max_column
            
            # Ajustar el ancho de las columnas
            ws.column_dimensions[get_column_letter(1)].width = 40 # Columna de cuenta más ancha

            # --- 2.1 Replicar Encabezados y Títulos (Fila 1 a Fila 3) ---

            # Título: ESTADO DE SITUACION FINANCIERA (Fila 1)
            # Asumiendo que esto va en la Fila 1, celdas A1 a D1.
            if not ws['A1'].value: # Evita sobreescribir si ya hay contenido
                ws['A1'].value = "ESTADO DE SITUACION FINANCIERA"
                ws.merge_cells('A1:D1')
                ws['A1'].fill = HEADER_FILL
                ws['A1'].font = HEADER_FONT
                ws['A1'].alignment = Alignment(horizontal='left', vertical='center')

            # Títulos de columna (Cuenta, NOTA, 2020, 2019) (Fila 3)
            # Usando la imagen como referencia, los datos comienzan alrededor de la Fila 3.
            headers = ["Cuenta", "NOTA", "2020", "2019"]
            for col_idx, header in enumerate(headers, start=1):
                cell = ws.cell(row=3, column=col_idx)
                cell.value = header
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.border = THIN_BORDER
            
            # --- 2.2 Aplicar Bordes y Color de la Columna de Datos (Filas 4 en adelante) ---
            
            # Aplicar bordes al rango de datos de A4 hasta la última celda
            for row in range(4, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.border = THIN_BORDER
                    
                    # Estilo condicional: Aplicar color gris a las celdas de la columna 'Cuenta' (Columna A)
                    if col == 1:
                        cell.fill = SECTION_FILL
                        cell.font = Font(bold=True) # Los títulos principales en negrita
            
            # --- 2.3 Marcar la Celda Roja (B2 en la imagen, ajustada por la nueva estructura) ---
            # Si se decide poner la marca roja, hay que reubicarla si ajustamos las filas
            try:
                # El marcador rojo de la nota de activos puede estar en la celda B4 o B5
                ROJO = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                if celda_marcada in ws:
                    ws[celda_marcada].value = "●"
                    ws[celda_marcada].fill = ROJO
            except:
                pass

            # 3. Guardar el archivo con estilos
            wb.save(path_xlsx)
            resultados.append(path_xlsx)

    return resultados
