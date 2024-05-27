import openpyxl
from openpyxl import Workbook
from Pila import Pila

class GeneradorDeArchivos:

    # Genera los archivos csv tomados del Excel. Recibe la direccion del archivo Excel y la ruta donde se guardarán los archivos.

    def __init__(self, rutaWb, rutaArchivos, colorInicial, colorFinal):
        self.rutaWb = rutaWb
        self.rutaArchivos = rutaArchivos
        self.colorIncial = colorInicial
        self.colorFinal = colorFinal

    def generarArchivos(self):
        wb = openpyxl.load_workbook(self.rutaWb, data_only=True)
        ws = wb.active
        COLUM_MAX = 27
        header = ('Cliente;Descripcion;D.Compensacion;N.Recibo;F.Contabilizacion;Nro.Cheque;;Imp.Valores;Imp.Documento;Descripcion;'
         'Nro. Recibo/Factura;Factura SAP-Nro Cbte.;Banco;Descripcion;Dias;F.Base;F.Vencimiento;Imp.Aplicado;Imp.Documento;'
         'Saldo;Atraso;Numerales;Dias Pago;Numerales Pago;Intereses;Moneda;Cambio')
        
        celdaIni = 0

        for fila in ws.iter_rows(max_col = 30):
            for celda in fila:
                if(celda.fill.fgColor.rgb == self.colorIncial and celdaIni == 0):
                    celdaIni = celda
                    fp = open(self.rutaArchivos + '\\' + celdaIni.value + ".csv", "wt")
                    fp.write(header + '\n')
                    
                #Tengo que revisar que onda como detectar la celda Y*ROW* sea mayor a 0, falta eso

            
                    
                if(celda.column_letter == 'Y' and celda.fill.fgColor.rgb == self.colorFinal and celda.value < 0 ):
                    i = celdaIni.row
                    print(ws['Y'+ str(celda.row)])
                    while i < celda.row:
                        reemplazar = ['None', '[', ']', '\'', '\"']
                        data = str([ws.cell(row=i, column = i).value for i in range(1, COLUM_MAX + 1)])

                        for caracter in reemplazar:
                            data = data.replace(caracter, '')

                        data = data.replace(',', ';')
                        fp.write(data + '\n')
                        i += 1

                    celdaIni = 0
                    fp.close()

