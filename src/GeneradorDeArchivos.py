import openpyxl
import os
import io

# Genera los archivos csv tomados del Excel. Recibe la direccion del archivo Excel y la ruta donde se guardarán los archivos.

class GeneradorDeArchivos:
    def __init__(self, rutaWb, rutaArchivos, colorInicial, colorFinal):
        self.rutaWb = rutaWb
        self.rutaArchivos = rutaArchivos
        self.colorIncial = colorInicial
        self.colorFinal = colorFinal

    def generarArchivos(self):
        with open(self.rutaWb, "rb") as f:
            in_mem_file = io.BytesIO(f.read())

        wb = openpyxl.load_workbook(in_mem_file, data_only=True, read_only=False)
        ws = wb.active
        COLUM_MAX = 27
        header = ('Cliente;Descripcion;D.Compensacion;N.Recibo;F.Contabilizacion;Nro.Cheque;;Imp.Valores;Imp.Documento;Descripcion;'
         'Nro. Recibo/Factura;Factura SAP-Nro Cbte.;Banco;Descripcion;Dias;F.Base;F.Vencimiento;Imp.Aplicado;Imp.Documento;'
         'Saldo;Atraso;Numerales;Dias Pago;Numerales Pago;Intereses;Moneda;Cambio')
        
        celdaIni = 0

        for fila in ws.iter_rows(max_col = COLUM_MAX + 1):
            fp = 0
            for celda in fila:
                if(celda.fill.fgColor.rgb == self.colorIncial and celdaIni == 0):
                    celdaIni = celda
 
                if(celda.column_letter == 'Y' and celda.fill.fgColor.rgb == self.colorFinal):
                    if(celda.value < 0):
                      i = celdaIni.row
                      fp = open(self.rutaArchivos + '\\' + celdaIni.value + ".csv", "wt")
                      fp.write(header + '\n')

                      while i < celda.row + 1:
                          reemplazar = ['None', '[', ']', '\'', '\"']
                          data = str([ws.cell(row=i, column = j).value for j in range(1, COLUM_MAX + 1)])
 
                          for caracter in reemplazar:
                              data = data.replace(caracter, '')

                          data = data.replace(',', ';')
                          fp.write(data + '\n')

                          i += 1

                    celdaIni = 0
                    if (fp != 0):
                        fp.close()
        wb.close()