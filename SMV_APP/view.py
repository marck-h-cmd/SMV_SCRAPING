from django.http import HttpResponse, JsonResponse
from django.template import Template, Context
from django.template.loader import get_template
from django.shortcuts import render
from .scraper import ejecutar_scraping_smv
from django.views.decorators.csrf import csrf_exempt
import json
import os
import pandas as pd
from pathlib import Path
import mimetypes
from django.http import FileResponse
import logging
from django.views.decorators.http import require_http_methods
from SMV_APP.analisis import formato_xls_xlsx, union_archivos, analisis_VH, analisis_Ratios, graficosRatios, renombrar

logger = logging.getLogger(__name__)

@csrf_exempt
def descargar_datos_financieros(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            empresa_nombre = data.get('empresa_nombre', '')

            if not empresa_nombre:
                return JsonResponse({'error': 'Nombre de empresa requerido'})
            
    
            resultado = ejecutar_scraping_smv(empresa_nombre, 2024,5)
            
            return JsonResponse(resultado)
            
        except Exception as e:
            logger.error(f"Error en descarga de datos: {str(e)}")
            return JsonResponse({'error': str(e)})
    
    return JsonResponse({'error': 'Método no permitido'})

@csrf_exempt
def verificar_archivos(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            path = data.get('path', '')
            
            if not path:
                return JsonResponse({'error': 'Ruta requerida'})
            
          
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            full_path = os.path.join(base_dir, path)
            print("Verificando ruta:", full_path)
            
            archivos = []
            
            if os.path.exists(full_path) and os.path.isdir(full_path):
                for filename in os.listdir(full_path):
                    file_path = os.path.join(full_path, filename)
                    
          
                    if filename.lower().endswith(('.xls', '.xlsx')):
                        try:
                            stat = os.stat(file_path)
                            
                            year = None
                            for part in filename.split('-'):
                                if part.isdigit() and len(part) == 4:
                                    year = int(part)
                                    break
                            
                            archivo_info = {
                                'nombre': filename,
                                'ruta': os.path.join(path, filename).replace('\\', '/'),
                                'tamaño': stat.st_size,
                                'fecha': stat.st_mtime * 1000, 
                                'tipo': 'excel',
                                'año': year
                            }
                            archivos.append(archivo_info)
                            
                        except Exception as e:
                            logger.warning(f"Error procesando archivo {filename}: {str(e)}")
                            continue
            
            return JsonResponse({
                'archivos': archivos,
                'total': len(archivos)
            })
            
        except Exception as e:
            logger.error(f"Error verificando archivos: {str(e)}")
            return JsonResponse({'error': str(e)})
    
    return JsonResponse({'error': 'Método no permitido'})

@csrf_exempt
def preview_excel(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            file_path = data.get('file_path', '')
            
            if not file_path:
                return JsonResponse({'error': 'Ruta del archivo requerida'})
            
       
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            full_path = os.path.join(base_dir, file_path)
            
            if not os.path.exists(full_path):
                return JsonResponse({'error': 'Archivo no encontrado'})
            
            try:
            
                if file_path.lower().endswith('.xlsx'):
                    df = pd.read_excel(full_path, engine='openpyxl', nrows=20)  
                else:  
                    df = pd.read_excel(full_path, engine='xlrd', nrows=20)  
                
                
                data_list = []
                
                
                headers = df.columns.tolist()
                data_list.append(headers)
                
                
                for _, row in df.iterrows():
                    row_data = []
                    for value in row:
                        
                        if pd.isna(value):
                            row_data.append('')
                        else:
                            row_data.append(str(value))
                    data_list.append(row_data)
                
                return JsonResponse({
                    'data': data_list,
                    'total_rows': len(df),
                    'total_columns': len(df.columns),
                    'filename': os.path.basename(file_path)
                })
                
            except Exception as e:
                logger.error(f"Error leyendo archivo Excel {file_path}: {str(e)}")
                return JsonResponse({'error': f'Error leyendo archivo Excel: {str(e)}'})
            
        except Exception as e:
            logger.error(f"Error en vista previa: {str(e)}")
            return JsonResponse({'error': str(e)})
    
    return JsonResponse({'error': 'Método no permitido'})

@csrf_exempt
def download_file(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            file_path = data.get('file_path', '')
            
            if not file_path:
                return JsonResponse({'error': 'Ruta del archivo requerida'})
   
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            full_path = os.path.join(base_dir, file_path)
            
            if not os.path.exists(full_path):
                return JsonResponse({'error': 'Archivo no encontrado'})
            
 
            content_type, _ = mimetypes.guess_type(full_path)
            if content_type is None:
                content_type = 'application/octet-stream'
            
         
            try:
                response = FileResponse(
                    open(full_path, 'rb'),
                    content_type=content_type,
                    as_attachment=True,
                    filename=os.path.basename(file_path)
                )
                
                return response
                
            except Exception as e:
                logger.error(f"Error sirviendo archivo {file_path}: {str(e)}")
                return JsonResponse({'error': f'Error sirviendo archivo: {str(e)}'})
            
        except Exception as e:
            logger.error(f"Error en descarga: {str(e)}")
            return JsonResponse({'error': str(e)})
    
    return JsonResponse({'error': 'Método no permitido'})

@csrf_exempt
def delete_file(request):
   
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            file_path = data.get('file_path', '')
            
            if not file_path:
                return JsonResponse({'error': 'Ruta del archivo requerida'})
            
       
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            full_path = os.path.join(base_dir, file_path)
            
            if not os.path.exists(full_path):
                return JsonResponse({'error': 'Archivo no encontrado'})
            
            try:
             
                os.remove(full_path)
                
             
                parent_dir = os.path.dirname(full_path)
                if os.path.exists(parent_dir) and not os.listdir(parent_dir):
                    try:
                        os.rmdir(parent_dir)
                        logger.info(f"Carpeta vacía eliminada: {parent_dir}")
                    except Exception as e:
                        logger.warning(f"No se pudo eliminar carpeta vacía {parent_dir}: {str(e)}")
                
                return JsonResponse({
                    'success': True,
                    'message': 'Archivo eliminado correctamente',
                    'deleted_file': os.path.basename(file_path)
                })
                
            except Exception as e:
                logger.error(f"Error eliminando archivo {file_path}: {str(e)}")
                return JsonResponse({'error': f'Error eliminando archivo: {str(e)}'})
            
        except Exception as e:
            logger.error(f"Error en eliminación: {str(e)}")
            return JsonResponse({'error': str(e)})
    
    return JsonResponse({'error': 'Método no permitido'})

def acceder(request):
    return render(request,"main/index.html")


@csrf_exempt
@require_http_methods(["POST"])
def analisis(request):
    try:
        # 1. Obtener la ruta de la empresa del JSON del POST
        data = json.loads(request.body)
        
        # Asumiendo que el frontend envía la ruta de la carpeta (ej. 'descargas_smv/ENERGIA_DEL_PACIFICO_SA')
        carpeta_empresa = data.get('carpeta_empresa', '') 

        empresa_clean = "".join(c for c in carpeta_empresa if c.isalnum() or c in (' ', '-', '_')).rstrip()
        empresa_clean = empresa_clean.replace(' ', '_')
        
        if not empresa_clean:
             return JsonResponse({'error': 'La ruta de la carpeta de la empresa es requerida.'}, status=400)

        # 2. Definir la base para las rutas (similar a tus otras funciones)
        # Esto asume que el path es relativo a la raíz del proyecto
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        
        # 3. Construir las rutas dinámicamente
        
        # La ruta base donde están todos los archivos .xls o .xlsx
        DIR_BASE_FINANCIEROS = os.path.join("descargas_smv", empresa_clean)

        formato_xls_xlsx(DIR_BASE_FINANCIEROS)
        
        # Asumimos que los archivos tienen nombres predecibles basados en el año y el directorio:
        RUTA1 = os.path.join(DIR_BASE_FINANCIEROS, "2024-ReporteDetalleInformacionFinanciero.xlsx")
        RUTA2 = os.path.join(DIR_BASE_FINANCIEROS, "2023-ReporteDetalleInformacionFinanciero.xlsx")
        RUTA3 = os.path.join(DIR_BASE_FINANCIEROS, "2022-ReporteDetalleInformacionFinanciero.xlsx")
        RUTA4 = os.path.join(DIR_BASE_FINANCIEROS, "2021-ReporteDetalleInformacionFinanciero.xlsx")
        RUTA5 = os.path.join(DIR_BASE_FINANCIEROS, "2020-ReporteDetalleInformacionFinanciero.xlsx")
        
        # Verificar que la ruta principal (RUTA1) exista antes de proceder
        if not os.path.exists(RUTA1):
             return JsonResponse({'error': f"El archivo base (2024) no fue encontrado en: {RUTA1}"}, status=404)
        
        # 4. Ejecutar formato, unión y análisis
        union_archivos(RUTA2, RUTA1, 5)
        union_archivos(RUTA3, RUTA1, 6)
        union_archivos(RUTA4, RUTA1, 7)
        union_archivos(RUTA5, RUTA1, 8)

        analisis_VH(RUTA1)
        analisis_Ratios(RUTA1)
        graficosRatios(RUTA1)
        renombrar(RUTA1)

        return JsonResponse({
            "status": "success", 
            "message": "Análisis y unión completados con éxito."
        }, status=200)

    except Exception as e:
        logger.error(f"Error en el proceso de análisis y unión: {str(e)}")
        return JsonResponse({
            "status": "error", 
            "message": f"Error interno en la función de análisis: {str(e)}"
        }, status=500)