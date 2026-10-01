import glob
import os
import configparser
import xlwings as xw

def load_config():
    config = configparser.ConfigParser()
    if not os.path.exists("config.conf"):
        raise FileNotFoundError("File 'config.conf' tidak ditemukan! Pastikan file berada di folder yang sama.")
    config.read("config.conf")
    return config["WINGS"]

def process_data():
    wings_config = load_config()
    irc_del = wings_config.get("irc-del", "2:19593")
    irc_hid = wings_config.get("irc-hid", "D:G")
    zn_del = wings_config.get("zn-del", "2:75")
    zn_hid = wings_config.get("zn-hid", "D:H")

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False

    try:
        for file_path in glob.glob("ORDER IRC JATENG*.xlsx"):
            print(f"--> Memproses {file_path}")
            abs_path = os.path.abspath(file_path)
            
            wb = app.books.open(abs_path, corrupt_load=1)
            ws = wb.sheets[0]
            
            ws.api.AutoFilterMode = False
            
            ws.range(irc_del).api.Delete()
            ws.range(irc_hid).api.EntireColumn.Hidden = True
            ws.range('C:C').column_width = 30
            
            lr = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
            last_row = lr if lr > 1 else 2
            
            ws.range(f'A1:I{last_row}').api.AutoFilter(Field=9, Criteria1="=")
            wb.save(abs_path)
            wb.close()

        for file_path in glob.glob("ORDER ZN JATENG*.xlsx"):
            print(f"--> Memproses {file_path}")
            abs_path = os.path.abspath(file_path)
            
            wb = app.books.open(abs_path, corrupt_load=1)
            ws = wb.sheets[0]
            
            ws.api.AutoFilterMode = False
            
            ws.range(zn_del).api.Delete()
            ws.range(zn_hid).api.EntireColumn.Hidden = True
            ws.range('C:C').column_width = 30
            
            lr = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
            last_row = lr if lr > 1 else 2
            
            ws.range(f'A1:I{last_row}').api.AutoFilter(Field=9, Criteria1="=")
            wb.save(abs_path)
            wb.close()
            
    finally:
        app.quit()

if __name__ == "__main__":
    process_data()
    print("--> Selesai")