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
from SMV_APP.analisis import formato_xls_xlsx, union_archivos, analisis_Ratios, graficosRatios, renombrar

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


@require_http_methods(["POST"])
def analisis(request):
    # 1. Ejecutar formato de excels 
    formato_xls_xlsx()

    RUTA5 = r"C:\Users\FELIX\Desktop\CICLO 6\FINANZAS CORPORATIVAS\PROYECTO\SMV_SCRAPING\descargas_smv\ENERGIA_DEL_PACIFICO_SA\2020-ReporteDetalleInformacionFinanciero.xlsx"
    RUTA4 = r"C:\Users\FELIX\Desktop\CICLO 6\FINANZAS CORPORATIVAS\PROYECTO\SMV_SCRAPING\descargas_smv\ENERGIA_DEL_PACIFICO_SA\2020-ReporteDetalleInformacionFinanciero.xlsx"
    RUTA3 = r"C:\Users\FELIX\Desktop\CICLO 6\FINANZAS CORPORATIVAS\PROYECTO\SMV_SCRAPING\descargas_smv\ENERGIA_DEL_PACIFICO_SA\2020-ReporteDetalleInformacionFinanciero.xlsx"
    RUTA2 = r"C:\Users\FELIX\Desktop\CICLO 6\FINANZAS CORPORATIVAS\PROYECTO\SMV_SCRAPING\descargas_smv\ENERGIA_DEL_PACIFICO_SA\2020-ReporteDetalleInformacionFinanciero.xlsx"
    RUTA1 = r"C:\Users\FELIX\Desktop\CICLO 6\FINANZAS CORPORATIVAS\PROYECTO\SMV_SCRAPING\descargas_smv\ENERGIA_DEL_PACIFICO_SA\2024-ReporteDetalleInformacionFinanciero.xlsx"
    
    try:
        # Ejecución de lógica
        union_archivos(RUTA2, RUTA1, 6)
        union_archivos(RUTA3, RUTA1, 7)
        union_archivos(RUTA4, RUTA1, 8)
        union_archivos(RUTA5, RUTA1, 9)
        analisis_VH(RUTA1)
        analisis_Ratios(RUTA1)
        graficosRatios(RUTA1)
        renombrar(RUTA1)
        return JsonResponse({
            "status": "success", 
            "message": "Análisis y unión completados con éxito."
        }, status=200)

    except Exception as e:
        print(f"Error en el proceso de análisis y unión: {e}")
        return JsonResponse({"status": "error", "message": f"Error interno en la función de análisis: {str(e)}"}, status=500)