from django.http import HttpResponse
from django.template import Template,Context
from django.template.loader import get_template
from django.shortcuts import render
from scraper import ejecutar_scraping_smv
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
import json

@csrf_exempt
def descargar_datos_financieros(request):
    if request.method == 'POST':
        try:
            data = json.loads(request.body)
            empresa_nombre = data.get('empresa_nombre', '')
            anios = data.get('anios', [2024, 2022, 2020])
            
            if not empresa_nombre:
                return JsonResponse({'error': 'Nombre de empresa requerido'})
            
            # Ejecutar scraping
            resultado = ejecutar_scraping_smv(empresa_nombre, anios)
            
            return JsonResponse(resultado)
            
        except Exception as e:
            return JsonResponse({'error': str(e)})
    
    return JsonResponse({'error': 'Método no permitido'})

def acceder(request):
    return render(request,"main/index.html")
