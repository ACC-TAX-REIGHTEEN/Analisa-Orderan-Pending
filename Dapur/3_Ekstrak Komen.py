import os
import glob
import re
import sys
import configparser
import openpyxl
import pandas as pd
from datetime import datetime

CONFIG_FILE = "config.conf"

BULAN_INDO = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MEI", 6: "JUN",
    7: "JUL", 8: "AGU", 9: "SEP", 10: "OKT", 11: "NOV", 12: "DES"
}

PATTERN_RINCIAN = re.compile(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+([\d\.]+)')

def format_tanggal_custom(tgl_str):
    if not tgl_str or str(tgl_str).lower() == "none":
        return ""
    try:
        tgl_clean = str(tgl_str).split()[0]
        dt = datetime.strptime(tgl_clean, "%Y-%m-%d")
        
        day = f"{dt.day:02d}"
        month = BULAN_INDO[dt.month]
        year = f"{dt.year % 100:02d}" 
        
        return f"{day} {month} {year}"
    except Exception:
        return str(tgl_str)

def ekstrak_data_order(file_pattern, col_idx, output_file, label_proses):
    print(f"--> Memproses ekstraksi data: {label_proses}")

    files = glob.glob(file_pattern)
    if not files:
        print(f"--> Lewati file Excel dengan pola '{file_pattern}' tidak ditemukan.")
        return

    file_path = files[0]
    print(f"--> Membuka file: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    extracted_data = []

    for row in sheet.iter_rows(min_row=2):
        row_idx = row[0].row
        
        if sheet.row_dimensions[row_idx].hidden:
            continue

        tgl_input_raw = str(row[0].value) if row[0].value is not None else ""
        nama = str(row[1].value) if row[1].value is not None else ""
        toko = str(row[2].value) if row[2].value is not None else ""
        keterangan = str(row[9].value) if row[9].value is not None else ""
        
        if not tgl_input_raw and not nama and not toko:
            continue

        tgl_input = format_tanggal_custom(tgl_input_raw)

        cell_target = row[col_idx]
        total_kolom_k = 0.0
        
        if cell_target.value is not None:
            val_str = str(cell_target.value).split('\n')[0]
            val_clean = re.sub(r'[^\d]', '', val_str)
            if val_clean:
                total_kolom_k = float(val_clean)
                
        rincian_text = ""
        if cell_target.comment is not None:
            rincian_text += str(cell_target.comment.text) + " "
        
        if isinstance(cell_target.value, str):
            rincian_text += str(cell_target.value)

        matches = PATTERN_RINCIAN.findall(rincian_text)
        
        if matches:
            for match in matches:
                tgl_nota = match[0]
                tgl_jth_tempo = match[1]
                nominal_str = re.sub(r'[^\d]', '', match[2])
                nominal_rincian = float(nominal_str) if nominal_str else 0.0
                
                extracted_data.append({
                    "Tgl Input": tgl_input,
                    "Nama": nama,
                    "Toko": toko,
                    "Total Kolom K": total_kolom_k,
                    "Tgl Nota": tgl_nota,
                    "Tgl Jth Tempo": tgl_jth_tempo,
                    "Nominal Rincian": nominal_rincian,
                    "Status": "OK" if nominal_rincian > 0 else "Check Format Angka",
                    "Keterangan": keterangan
                })
        else:
            extracted_data.append({
                "Tgl Input": tgl_input,
                "Nama": nama,
                "Toko": toko,
                "Total Kolom K": total_kolom_k,
                "Tgl Nota": "",
                "Tgl Jth Tempo": "",
                "Nominal Rincian": 0.0,
                "Status": "Tidak Ada Rincian",
                "Keterangan": keterangan
            })

    df_result = pd.DataFrame(extracted_data)
    df_result.to_excel(output_file, index=False)
    print(f"--> Berhasil! Data {label_proses} disimpan ke: {output_file}")

def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"--> Error file '{CONFIG_FILE}' tidak ditemukan!")
        return

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)

    url_irc = ""
    url_zn = ""

    if config.has_section('SS'):
        url_irc = config.get('SS', 'url-irc', fallback='').strip()
        url_zn = config.get('SS', 'url-zn', fallback='').strip()
    else:
        print("--> Peringatan section [SS] tidak ditemukan di config.conf!")

    if url_irc:
        ekstrak_data_order(
            file_pattern="ORDER IRC JATENG*.xlsx",
            col_idx=7,
            output_file="Hasil_Ekstrak_Rincian_IRC_temp.xlsx",
            label_proses="IRC"
        )
    else:
        print("\n--> Lewati 'url-irc' di config.conf kosong. Proses ekstraksi IRC dilewati.")

    if url_zn:
        ekstrak_data_order(
            file_pattern="ORDER ZN JATENG*.xlsx",
            col_idx=10,
            output_file="Hasil_Ekstrak_Rincian_ZN_temp.xlsx",
            label_proses="ZN"
        )
    else:
        print("\n--> Lewati 'url-zn' di config.conf kosong. Proses ekstraksi ZN dilewati.")

    print("\n--> Seluruh proses ekstraksi selesai.")

if __name__ == "__main__":
    main()