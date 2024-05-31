from GeneradorDeArchivos import GeneradorDeArchivos
from Conversor import Conversor
from os import scandir, remove


rutaExcel = "C:\\Users\\nataniel.cinquegrani\\OneDrive - hzgroup.com.ar\\Desktop\\excel.XLSX"
rutaArchivos = "C:\\Users\\nataniel.cinquegrani\\OneDrive - hzgroup.com.ar\\Desktop\\Tablas"
colorNaranja = "FFFFCC00"
colorAmarillo = "FFFFFF00"

def main():
    generador = GeneradorDeArchivos(rutaExcel, rutaArchivos, colorNaranja, colorAmarillo)
    generador.generarArchivos()
    conversor = Conversor
    
    for arch in scandir(rutaArchivos):
      if arch.is_file():
        rutaConvertido = rutaArchivos + arch.name + ".xlsx"
        rutaConvertido = rutaConvertido.replace(".csv", "")
        conversor.convertir_csv_xlsx(arch.path, rutaConvertido)
        remove(arch)
      

main()
