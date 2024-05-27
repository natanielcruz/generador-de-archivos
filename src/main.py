from GeneradorDeArchivos import GeneradorDeArchivos


rutaExcel = "C:\\Users\\nataniel.cinquegrani\\OneDrive - hzgroup.com.ar\\Desktop\\excel.XLSX"
rutaArchivos = "C:\\Users\\nataniel.cinquegrani\\OneDrive - hzgroup.com.ar\\Desktop\\Tablas"
colorNaranja = "FFFFCC00"
colorAmarillo = "FFFFFF00"

def main():
    generador = GeneradorDeArchivos(rutaExcel, rutaArchivos, colorNaranja, colorAmarillo)
    generador.generarArchivos()


main()
