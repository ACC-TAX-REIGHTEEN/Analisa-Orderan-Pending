import glob
import os
import xlwings as xw

def process_data():
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False

    for file_path in glob.glob("ORDER IRC JATENG*.xlsx"):
        print(f"--> Memproses {file_path}")
        abs_path = os.path.abspath(file_path)
        wb = app.books.open(abs_path)
        ws = wb.sheets[0]
        
        ws.api.AutoFilterMode = False
        ws.range('2:19593').api.Delete()
        ws.range('D:G').api.EntireColumn.Hidden = True
        ws.range('C:C').column_width = 30
        
        lr = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
        last_row = lr if lr > 1 else 2
        
        ws.range(f'A1:I{last_row}').api.AutoFilter(Field=9, Criteria1="=")
        wb.save()
        wb.close()

    for file_path in glob.glob("ORDER ZN JATENG*.xlsx"):
        print(f"--> Memproses {file_path}")
        abs_path = os.path.abspath(file_path)
        wb = app.books.open(abs_path)
        ws = wb.sheets[0]
        
        ws.api.AutoFilterMode = False
        ws.range('2:75').api.Delete()
        ws.range('D:H').api.EntireColumn.Hidden = True
        ws.range('C:C').column_width = 30
        
        lr = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
        last_row = lr if lr > 1 else 2
        
        ws.range(f'A1:I{last_row}').api.AutoFilter(Field=9, Criteria1="=")
        wb.save()
        wb.close()

if __name__ == "__main__":
    process_data()
    print("--> Selesai")