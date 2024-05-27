from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Protection
import csv

class Conversor:
    def __init__(self, rutaArchivo, rutaConvertido):
        self.rutaArchivo = rutaArchivo
        self.rutaConvertido = rutaConvertido
        
    def convertir_csv_xlsx(self):
        wb = Workbook()
        ws = wb.active

        with open(self.rutaArchivo, 'rt') as fp:
            for row in csv.reader(fp, delimiter=';'):
                ws.append(row)

        for rows in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1):
            for cell in rows:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="00c1f5", end_color="00c1f5", fill_type="solid")
                cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))
        
        wb.save(self.rutaConvertido)
