from GeneradorDeArchivos import GeneradorDeArchivos
from Conversor import Conversor
from datetime import datetime
from os import scandir, remove
import json
import os
import sys

# Obtiene la ruta de la configuración tanto para el ejecutable como para el archivo .py.

def cargar_configuracion(nombreArchivo):      
   with open (os.path.join(                                 # Hace un join de la ruta completa.
      os.path.dirname(sys.executable)                       # No sube de nivel el directorio, por lo que el ejecutable debe estar en la misma ubicacion que config.
      if getattr(sys, "frozen", False)                      # Detecta si se esta ejecutando el .exe o el .py.
      else os.path.join(os.path.dirname(__file__), ".."),   # Sube un directorio la ruta del archivo main.py (equivale a "../src").
      'config', nombreArchivo), "rt") as fp:                # Abre el archivo como texto en formato solo lectura.
      return json.load(fp) 


# El programa propiamente dicho

def main():
    
    fechaYHora = datetime.now()

    rutaConfiguracion = str("config.json")               # Se carga el archivo de configuración para poder obtener los datos.
    config = cargar_configuracion(rutaConfiguracion)
    
    rutaExcel = config["rutaExcel"]                      # Ruta del archivo Excel a procesar.
    rutaExcelProcesado = config["rutaExcelProcesado"]    # Ruta donde se moverá el archivo Excel una vez procesado.
    rutaArchivos = config["rutaArchivos"]                # Ruta donde se almacenarán los archivos a enviar generados.
    naranja = config["primerColor"]                      # Color utilizado para saber donde comienza la tabla de un cliente en el archivo Excel.
    amarillo = config["segundoColor"]                    # Color utilizado para saber donde finaliza la tabla de un cliente en el archivo Excel.

    generador = GeneradorDeArchivos(rutaExcel, rutaArchivos, naranja, amarillo) # Genera los archivos en formato .csv.
    generador.generarArchivos()
    conversor = Conversor
    
    for arch in scandir(rutaArchivos):                               # Transforma los archivos en formato .csv a .xlsx. Además, personaliza la apariencia (colores y bordes).
      if arch.is_file() and arch.path.__contains__(".csv"):
        arch.path
        rutaConvertido = rutaArchivos + "\\" + arch.name + ".xlsx"
        rutaConvertido = rutaConvertido.replace(".csv", "")
        conversor.convertir_csv_xlsx(arch.path, rutaConvertido)
        remove(arch)

    os.rename(rutaExcel, rutaExcelProcesado + '\\Procesado ' + fechaYHora.strftime("%d-%m-%Y %H%M") +'.xlsx') # Cambia el nombre del archivo Excel para saber que ya fue procesado.


# Llama a la función main

if __name__ == "__main__":
   main()

# $pyinstaller --onefile --add-data "config/config.json;config" "src/main.py"    -> Este comando sirve para transformar el programa a un ejecutable .exe y que pueda ser trasladado a otra computadora.
#                                                                                -> Debe ejecutarse en una consola bash y tener instalado el paquete pyinstaller.