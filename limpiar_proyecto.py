#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SCRIPT DE LIMPIEZA PARA EL PROYECTO
===================================

Limpia archivos temporales y mantiene el proyecto organizado.

Uso: python limpiar_proyecto.py
"""

import shutil
from pathlib import Path

def limpiar_proyecto():
    """Limpiar archivos y directorios temporales del proyecto"""
    
    project_root = Path(__file__).parent
    
    # Directorios a limpiar
    dirs_to_clean = [
        "RESULTADOS_TEST_FINAL",
        "test_output",  
        "test_output_complete",
        "test_individual",
        "__pycache__",
        "SMV_APP/__pycache__"
    ]
    
    # Archivos a limpiar
    files_to_clean = [
        "*.pyc",
        "*.pyo", 
        "*.log",
        "db.sqlite3"
    ]
    
    print("🧹 LIMPIANDO PROYECTO...")
    print("=" * 30)
    
    # Limpiar directorios
    for dir_name in dirs_to_clean:
        dir_path = project_root / dir_name
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"🗑️  Eliminado: {dir_name}/")
            
    # Limpiar archivos por patrón
    for pattern in files_to_clean:
        for file_path in project_root.rglob(pattern):
            if file_path.is_file():
                file_path.unlink()
                print(f"🗑️  Eliminado: {file_path.name}")
                
    print("\n✅ ¡Proyecto limpio y organizado!")
    print("\n📁 Archivos principales conservados:")
    print("   • test_definitivo.py (Test principal)")
    print("   • SMV_APP/ (Código del sistema)")
    print("   • descargas_smv/ (Datos originales)")
    print("   • manage.py (Django)")

if __name__ == "__main__":
    limpiar_proyecto()