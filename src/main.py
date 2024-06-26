from GeneradorDeArchivos import GeneradorDeArchivos
from Conversor import Conversor
from datetime import datetime
from os import scandir, remove
import json
import os
import sys

# Obtiene la ruta de la configuración tanto para el ejecutable como para el archivo .py

def cargar_configuracion(nombreArchivo):      
   with open (os.path.join(                                 # Hace un join de la ruta completa
      os.path.dirname(sys.executable)                       # No sube de nivel el directorio, por lo que el ejecutable debe estar en la misma ubicacion que config
      if getattr(sys, "frozen", False)                      # Detecta si se trata del ejecutable o de main.py 
      else os.path.join(os.path.dirname(__file__), ".."),   # Sube un directorio la ruta del archivo main.py "../src"
      'config', nombreArchivo), "rt") as fp:                # Abre el archivo como texto en formato solo lectura
      return json.load(fp) 

def main():
    
    fechaYHora = datetime.now()

    rutaConfiguracion = str("config.json")
    config = cargar_configuracion(rutaConfiguracion)
    
    rutaExcel = config["rutaExcel"]
    rutaExcelProcesado = config["rutaExcelProcesado"]
    rutaArchivos = config["rutaArchivos"]
    naranja = config["primerColor"]
    amarillo = config["segundoColor"]

    generador = GeneradorDeArchivos(rutaExcel, rutaArchivos, naranja, amarillo)
    generador.generarArchivos()
    conversor = Conversor
    
    for arch in scandir(rutaArchivos):
      if arch.is_file() and arch.path.__contains__(".csv"):
        arch.path
        rutaConvertido = rutaArchivos + "\\" + arch.name + ".xlsx"
        rutaConvertido = rutaConvertido.replace(".csv", "")
        conversor.convertir_csv_xlsx(arch.path, rutaConvertido)
        remove(arch)

    os.rename(rutaExcel, rutaExcelProcesado + '\\Procesado ' + fechaYHora.strftime("%d-%m-%Y %H%M") +'.xlsx') 
    rutaExcelProcesado
     
if __name__ == "__main__":
   main()

# $pyinstaller --onefile --add-data "config/config.json;config" "src/main.py"