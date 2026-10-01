import glob
import os
import configparser
import xlwings as xw

def load_config():
    config = configparser.ConfigParser()
    if not os.path.exists("config.conf"):
        raise FileNotFoundError("--> File 'config.conf' tidak ditemukan!")
    config.read("config.conf")
    return config["WINGS"]

def find_column_index(ws, target_header_name):
    if not target_header_name:
        return 9
    header_row = ws.range("1:1").value
    if header_row:
        target_clean = str(target_header_name).strip().lower()
        for idx, cell_value in enumerate(header_row, start=1):
            if cell_value is not None:
                cell_str = str(cell_value).strip().lower()
                if cell_str == target_clean:
                    return idx
    print(f"--> Peringatan: Header '{target_header_name}' tidak ditemukan. Menggunakan kolom 9.")
    return 9

def process_data():
    wings_config = load_config()
    
    irc_del = wings_config.get("irc-del", "2:19593")
    irc_hid = wings_config.get("irc-hid", "D:G")
    irc_filter_header = wings_config.get("irc-auto-filter", "Tgl Approved OWNER")

    zn_del = wings_config.get("zn-del", "2:75")
    zn_hid = wings_config.get("zn-hid", "D:I")
    zn_filter_header = wings_config.get("zn-auto-filter", "Tgl Approved")

    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False

    try:
        tasks = [
            ("ORDER IRC JATENG*.xlsx", irc_del, irc_hid, irc_filter_header),
            ("ORDER ZN JATENG*.xlsx", zn_del, zn_hid, zn_filter_header)
        ]
        
        for pattern, del_range, hid_range, filter_header in tasks:
            for file_path in glob.glob(pattern):
                print(f"--> Memproses {file_path}")
                abs_path = os.path.abspath(file_path)
                wb = app.books.open(abs_path)
                ws = wb.sheets[0]
                
                ws.api.AutoFilterMode = False
                ws.range(del_range).api.Delete()
                ws.range(hid_range).api.EntireColumn.Hidden = True
                ws.range('C:C').column_width = 30
                
                lr = ws.range('A' + str(ws.cells.last_cell.row)).end('up').row
                last_row = lr if lr > 1 else 2
                
                filter_col = find_column_index(ws, filter_header)
                max_col = ws.range("1:1").end('right').column
                if max_col < filter_col:
                    max_col = filter_col
                
                ws.range((1, 1), (last_row, max_col)).api.AutoFilter(Field=filter_col, Criteria1="=")
                
                wb.save()
                wb.close()
    finally:
        app.quit()

if __name__ == "__main__":
    process_data()
    print("--> Selesai")