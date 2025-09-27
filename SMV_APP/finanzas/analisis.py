import os
import pandas as pd
from django.conf import settings

class FinancialAnalysis:
    def __init__(self, empresa):
        # Subimos un nivel desde BASE_DIR para llegar a "Programa"
        base_programa = os.path.dirname(settings.BASE_DIR)
        self.empresa = empresa
        self.folder_path = os.path.join(base_programa, "descargas_smv", empresa)
    def load_files(self):
        if not os.path.exists(self.folder_path):
            raise FileNotFoundError(f"No existe la carpeta: {self.folder_path}")
    
        archivos = [f for f in os.listdir(self.folder_path) if f.endswith(('.xls', '.xlsx'))]
    
        if not archivos:
            raise FileNotFoundError(f"No se encontraron archivos Excel en: {self.folder_path}")
    
        dataframes = []
        for archivo in archivos:
            ruta_archivo = os.path.join(self.folder_path, archivo)
            try:
                # Intentar como Excel real
                df = pd.read_excel(ruta_archivo, engine="xlrd")
            except Exception:
                # Si falla, probar como HTML disfrazado
                tables = pd.read_html(ruta_archivo)
                df = tables[0]  # la primera tabla
            dataframes.append(df)
    
        return dataframes

    def run_analysis(self):
        """Ejecuta el análisis financiero básico"""
        dataframes = self.load_files()

        if not dataframes:
            return {"error": "No se encontraron archivos financieros"}

        # Combinar todo en un solo DataFrame
        df_total = pd.concat(dataframes, ignore_index=True)

        # Guardar un Excel combinado
        output_file = os.path.join(self.folder_path, "analisis_completo.xlsx")
        df_total.to_excel(output_file, index=False)

        return {
            "mensaje": "Análisis generado correctamente",
            "archivos_procesados": len(dataframes),
            "archivo_salida": output_file,
        }

    
    def horizontal_analysis(self):
        """Ejemplo simple: variación año a año"""
        results = {}
        years = sorted(self.dataframes.keys())
        for i in range(1, len(years)):
            prev, curr = years[i-1], years[i]
            df = pd.DataFrame()
            df["Cuenta"] = self.dataframes[curr]["Cuenta"]
            df["Monto " + str(curr)] = self.dataframes[curr]["Monto"]
            df["Monto " + str(prev)] = self.dataframes[prev]["Monto"]
            df["Var%"] = (df["Monto " + str(curr)] - df["Monto " + str(prev)]) / df["Monto " + str(prev)] * 100
            results[f"{prev}-{curr}"] = df
        return results
    
    def vertical_analysis(self):
        """Ejemplo: cada cuenta como % del total"""
        results = {}
        for year, df in self.dataframes.items():
            total = df["Monto"].sum()
            df_copy = df.copy()
            df_copy["%"] = df_copy["Monto"] / total * 100
            results[year] = df_copy
        return results

    def export_excel(self, output_path, horizontal, vertical):
        """Guarda todo en un solo archivo Excel con varias hojas"""
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for year, df in self.dataframes.items():
                df.to_excel(writer, sheet_name=f"Original {year}", index=False)
            for k, df in horizontal.items():
                df.to_excel(writer, sheet_name=f"Horizontal {k}", index=False)
            for year, df in vertical.items():
                df.to_excel(writer, sheet_name=f"Vertical {year}", index=False)
        return output_path