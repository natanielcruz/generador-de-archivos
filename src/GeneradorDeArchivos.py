import openpyxl
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
        
        colorNaranja = 'FFFFCC00'
        colorAmarillo = 'FFFFFF00'

        celdas = Pila()
        celdas.crearPila()

        for fila in ws.iter_rows(max_col=1):
            for celda in fila:
                if(celda.fill.fgColor.rgb == self.colorIncial and (celdas.pilaVacia() == 1 or (celdas.verTope()).fill.fgColor.rgb != self.colorIncial)):
                    celdas.ponerEnPila(celda)
                    fp = open(self.rutaArchivos + '\\' + celda.value + ".csv", "wt")
                    fp.write(header + '\n')

                if(celdas.pilaVacia() == False):
                    reemplazar = ['None', '[', ']', '\'', '\"']
                    data = str([ws.cell(row=celda.row, column = i).value for i in range(1, COLUM_MAX + 1)])
                    for caracter in reemplazar:
                        data = data.replace(caracter, '')
                    data = data.replace(',', ';')
                    fp.write(data + '\n')
                    
                if(celda.fill.fgColor.rgb == self.colorFinal):
                    celdas.sacarDePila()
                    fp.close()

