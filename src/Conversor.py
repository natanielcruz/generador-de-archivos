from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill
import csv

class Conversor:    
    def convertir_csv_xlsx(rutaArchivo, rutaConvertido):
        wb = Workbook()
        ws = wb.active
        anchoCol = 14.5

        with open(rutaArchivo, 'rt') as fp:
            for row in csv.reader(fp, delimiter=';'):
                ws.append(row)
    
        for rows in ws.iter_rows(min_row=ws.max_row, max_row=ws.max_row, min_col=1):
           for cell in rows:
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="00c1f5", end_color="00c1f5", fill_type="solid")
                cell.border = Border(top=Side(style='medium'), bottom=Side(style='medium'))
                
        for col in ws.iter_cols():
            ws.column_dimensions[col[0].column_letter].width = anchoCol

        wb.save(rutaConvertido)
