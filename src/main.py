from GeneradorDeArchivos import GeneradorDeArchivos
from Conversor import Conversor


rutaExcel = "C:\\Users\\nataniel.cinquegrani\\OneDrive - hzgroup.com.ar\\Desktop\\excel.XLSX"
rutaArchivos = "C:\\Users\\nataniel.cinquegrani\\OneDrive - hzgroup.com.ar\\Desktop\\Tablas"
colorNaranja = "FFFFCC00"
colorAmarillo = "FFFFFF00"


def main():
   # generador = GeneradorDeArchivos(rutaExcel, rutaArchivos, colorNaranja, colorAmarillo)
  #  generador.generarArchivos()
    conversor = Conversor(rutaArchivos + "\\327.csv", rutaArchivos + "\\Excel\\prueba.xlsx")
    conversor.convertir_csv_xlsx()


main()
