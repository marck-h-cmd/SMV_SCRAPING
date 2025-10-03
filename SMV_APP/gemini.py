import os
import pandas as pd
import google.generativeai as genai
from dotenv import load_dotenv

load_dotenv()

class FinancialStatementAnalyzer:
    def __init__(self, api_key=None):
        self.api_key = api_key or os.getenv('GEMINI_API_KEY')
        if not self.api_key:
            raise ValueError("API key no encontrada")
        
        genai.configure(api_key=self.api_key)
        
        generation_config = {
            "temperature": 0.1,
            "top_p": 0.8,
            "top_k": 40,
        }
        
        safety_settings = [
            {
                "category": "HARM_CATEGORY_HARASSMENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_HATE_SPEECH",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
        ]
        
        self.model = genai.GenerativeModel(
            model_name="gemini-2.0-flash",
            generation_config=generation_config,
            safety_settings=safety_settings,
            system_instruction=self._get_financial_system_prompt()
        )
    
    def _get_financial_system_prompt(self):
        return """
        Eres un analista financiero senior especializado en análisis de estados financieros. 
        Tu objetivo es analizar exhaustivamente los estados financieros y responder específicamente segun los datos del documento proporcionado excel.:

        1. ¿Qué tendencias positivas o negativas se observan?
           - Identifica patrones de crecimiento o decrecimiento en ventas, utilidades, activos, pasivos
           - Evalúa tendencias en márgenes de rentabilidad
           - Analiza la evolución de la liquidez y solvencia

        2. ¿Qué ratios muestran fortaleza o debilidad?
           - Liquidez: corriente, rápida, prueba ácida
           - Endeudamiento: ratio deuda/patrimonio, deuda/activos
           - Rentabilidad: ROA, ROE, margen neto, margen operativo
           - Eficiencia: rotación de activos, inventarios, cuentas por cobrar

        3. ¿Cómo se relacionan los cambios en los ratios con el análisis vertical y horizontal ya realizado?
           - Explica cómo los cambios en ratios se correlacionan con variaciones porcentuales
           - Relaciona la composición porcentual (vertical) con la evolución temporal (horizontal)
           - Integra ambos análisis para dar una visión completa

        4. ¿Qué señales deberían llamar la atención de la gerencia?
           - Puntos críticos que requieren atención inmediata
           - Riesgos financieros identificados
           - Oportunidades de mejora detectadas
           - Recomendaciones accionables

        Usa terminología financiera precisa y parrafos cortos respondiendo.
        Estructura tu respuesta claramente en las cuatro secciones solicitadas solamente.
        """
    
    def read_excel_file(self, file_path):
        try:
            excel_file = pd.ExcelFile(file_path)
            sheets = {}
            
            for sheet_name in excel_file.sheet_names:
                sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
            
            return sheets
            
        except Exception as e:
            raise Exception(f"Error al leer el archivo Excel: {str(e)}")
    
    def describe_financial_data(self, df, sheet_name):
        description = f"HOJA: {sheet_name}\n"
        description += f"Dimensiones: {df.shape[0]} filas × {df.shape[1]} columnas\n"
        
        description += "Columnas financieras:\n"
        for col in df.columns:
            description += f"- {col}\n"
        
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            description += "\nResumen numérico:\n"
            for col in numeric_cols[:5]:
                description += f"- {col}: min={df[col].min():.2f}, max={df[col].max():.2f}, avg={df[col].mean():.2f}\n"
        
        return description
    
    def analyze_financial_statements(self, file_path):
        try:
            sheets = self.read_excel_file(file_path)
            
            if not sheets:
                return "No se encontraron hojas en el archivo Excel."
            
            data_context = f"ANÁLISIS DE ESTADOS FINANCIEROS: {os.path.basename(file_path)}\n\n"
            
            for sheet_name, df in sheets.items():
                data_context += self.describe_financial_data(df, sheet_name)
                data_context += "\n" + "-"*50 + "\n\n"
            
            analysis_prompt = f"""
            {data_context}
            
            Realiza un análisis financiero completo respondiendo estas preguntas específicas de acuerdo a los datos proporcionados en el documento:
            
            1. ¿Qué tendencias positivas o negativas se observan?
            2. ¿Qué ratios muestran fortaleza o debilidad?
            3. ¿Cómo se relacionan los cambios en los ratios con el análisis vertical y horizontal ya realizado?
            4. ¿Qué señales deberían llamar la atención de la gerencia?
            
            Proporciona un análisis detallado para cada punto.
            """
            
            response = self.model.generate_content(analysis_prompt)
            
            return response.text
            
        except Exception as e:
            return f"Error en el análisis financiero: {str(e)}"

def main():
    analyzer = FinancialStatementAnalyzer()
    
    excel_files = [r"C:\Users\FELIX\Desktop\CICLO 6\FINANZAS CORPORATIVAS\PROYECTO\SMV_SCRAPING\descargas_smv\ALICORP_SAA\ANALISIS\ANALISIS-ALICORP_SAA.xlsx"]    
    
    for file_path in excel_files:
        if os.path.exists(file_path):
            print(f"\n{'='*80}")
            print(f"ANALIZANDO: {file_path}")
            print(f"{'='*80}\n")
            
            analysis = analyzer.analyze_financial_statements(file_path)
            print(analysis)
            print(f"\n{'='*80}\n")
        else:
            print(f"Archivo no encontrado: {file_path}")

if __name__ == "__main__":
    main()