from GeneradorDeArchivos import GeneradorDeArchivos
from Conversor import Conversor
from os import scandir, remove
import json
import os
import sys

def cargar_configuracion(ruta):
   with open (os.path.join(os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(__file__), 'config', ruta), "rt") as fp:
      return json.load(fp)

def main():

    rutaConfiguracion = str("config.json")
    config = cargar_configuracion(rutaConfiguracion)

    rutaExcel = config["rutaExcel"]
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

if __name__ == "__main__":
   main()
