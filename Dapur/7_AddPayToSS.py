import os
import re
import configparser
import pandas as pd
import gspread
from gspread.utils import rowcol_to_a1

CONFIG_FILE = "config.conf"
CREDENTIALS_FILE = "credentials.json"

BULAN_INDO_REV = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MEI", 6: "JUN",
    7: "JUL", 8: "AGU", 9: "SEP", 10: "OKT", 11: "NOV", 12: "DES"
}

PATTERN_RINCIAN = re.compile(r'(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+([\d\.]+)')

def bersihkan_angka(val):
    if pd.isnull(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_clean = re.sub(r'[^\d]', '', str(val))
    return float(val_clean) if val_clean else 0.0

def normalisasi_tgl(tgl_str):
    if not tgl_str or str(tgl_str).lower() == "none":
        return ""
    
    s = str(tgl_str).strip().replace('\xa0', ' ')
    
    match_slash = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})', s)
    if match_slash:
        day = int(match_slash.group(1))
        month_num = int(match_slash.group(2))
        year = int(match_slash.group(3))
        
        if year < 100:
            year += 2000
            
        month_str = BULAN_INDO_REV.get(month_num, "")
        year_short = f"{year % 100:02d}"
        return f"{day:02d} {month_str} {year_short}".lower()
        
    return s.lower()

def bersihkan_prefix_lunas(teks_baris):
    return re.sub(r'^LUNAS\s*/\s*', '', str(teks_baris).strip(), flags=re.IGNORECASE).strip()

def ambil_semua_catatan(ws):
    try:
        res = ws.spreadsheet.fetch_sheet_metadata({
            'includeGridData': True,
            'ranges': [f"'{ws.title}'"],
            'fields': 'sheets(data(rowData(values(note))))'
        })
        
        notes_matrix = []
        sheets = res.get('sheets', [])
        if sheets and 'data' in sheets[0]:
            data = sheets[0]['data'][0]
            row_data = data.get('rowData', [])
            for r in row_data:
                row_notes = []
                values = r.get('values', [])
                for v in values:
                    note = v.get('note', '')
                    row_notes.append(note)
                notes_matrix.append(row_notes)
        return notes_matrix
    except Exception as e:
        print(f"--> Peringatan gagal mengambil catatan sel via API metadata: {e}")
        return []

def proses_pembaharuan_teks_catatan(teks_lama, rows_excel_toko):
    if not teks_lama:
        return teks_lama, False

    baris_list = str(teks_lama).replace('\xa0', ' ').split('\n')
    baris_baru = []
    ada_perubahan = False

    for baris in baris_list:
        baris_clean = baris.strip()
        if not baris_clean:
            baris_baru.append(baris)
            continue

        sudah_lunas = bool(re.match(r'^LUNAS\s*/', baris_clean, flags=re.IGNORECASE))

        if sudah_lunas:
            baris_baru.append(baris_clean)
            continue

        teks_murni = bersihkan_prefix_lunas(baris_clean)
        matches = PATTERN_RINCIAN.findall(teks_murni)

        baris_cocok = None
        if matches:
            tgl_nota_note, tgl_jt_note, nominal_str_note = matches[0]
            nominal_note = bersihkan_angka(nominal_str_note)

            for r_excel in rows_excel_toko:
                nominal_excel = bersihkan_angka(r_excel.get('Nominal Rincian', 0))
                no_faktur_excel = str(r_excel.get('No. Faktur', '')).strip()

                if abs(nominal_note - nominal_excel) < 0.01:
                    if no_faktur_excel and no_faktur_excel.lower() != 'nan':
                        baris_cocok = r_excel
                        break

        if baris_cocok:
            status_pelunasan = str(baris_cocok.get('Cek Pelunasan', '')).strip().lower()
            
            if status_pelunasan == 'lunas':
                teks_final = f"LUNAS/ {teks_murni}"
                ada_perubahan = True
                baris_baru.append(teks_final)
            else:
                baris_baru.append(baris_clean)
        else:
            baris_baru.append(baris_clean)

    teks_hasil = '\n'.join(baris_baru)
    return teks_hasil, ada_perubahan

def jalankan_writeback(url_sheet, sheet_name, col_target_name, file_excel_hasil, label_proses, gc):
    print(f"--> Write back Spreadsheets: {label_proses}")

    file_target = file_excel_hasil
    if not os.path.exists(file_target):
        file_temp = file_excel_hasil.replace('.xlsx', '_temp.xlsx')
        if os.path.exists(file_temp):
            file_target = file_temp
        else:
            print(f"--> Lewati file Excel hasil '{file_excel_hasil}' tidak ditemukan.")
            return

    print(f"--> Membaca data hasil analisis dari: {file_target}")
    df_excel = pd.read_excel(file_target)
    df_excel = df_excel.dropna(subset=['Tgl Input', 'Nama', 'Toko'], how='all')

    df_excel['tgl_input_norm'] = df_excel['Tgl Input'].apply(normalisasi_tgl)
    df_excel['nama_norm'] = df_excel['Nama'].astype(str).str.strip().str.lower()
    df_excel['toko_norm'] = df_excel['Toko'].astype(str).str.strip().str.lower()

    try:
        print(f"--> Menghubungkan ke Google Sheets {label_proses}...")
        sh = gc.open_by_url(url_sheet)
        ws = sh.worksheet(sheet_name)
    except Exception as e:
        print(f"--> Error gagal membuka Google Sheets: {e}")
        return

    all_values = ws.get_all_values()
    if not all_values:
        print("--> Lewati Google Sheets kosong.")
        return

    headers = [str(h).strip().lower() for h in all_values[0]]
    col_target_clean = str(col_target_name).strip().lower()

    target_col_idx = -1
    for idx, h in enumerate(headers):
        if h == col_target_clean:
            target_col_idx = idx
            break

    if target_col_idx == -1:
        target_col_idx = 7 if label_proses == "IRC" else 10
        print(f"--> Header '{col_target_name}' tidak ditemukan, menggunakan kolom indeks ke-{target_col_idx + 1}.")
    else:
        print(f"--> Kolom target '{col_target_name}' ditemukan pada kolom ke-{target_col_idx + 1}.")

    print("--> Mengambil catatan (cell notes) dari Google Sheets...")
    cell_notes = ambil_semua_catatan(ws)

    updates_notes = {}

    for r_idx in range(1, len(all_values)):
        row_gs = all_values[r_idx]
        if len(row_gs) <= target_col_idx:
            continue

        tgl_input_gs_raw = str(row_gs[0]).strip() if len(row_gs) > 0 else ""
        nama_gs = str(row_gs[1]).strip().lower() if len(row_gs) > 1 else ""
        toko_gs = str(row_gs[2]).strip().lower() if len(row_gs) > 2 else ""

        if not tgl_input_gs_raw and not nama_gs and not toko_gs:
            continue

        tgl_input_gs_norm = normalisasi_tgl(tgl_input_gs_raw)

        rows_excel_toko = df_excel[
            (df_excel['nama_norm'] == nama_gs) &
            (df_excel['toko_norm'] == toko_gs) &
            ((df_excel['tgl_input_norm'] == tgl_input_gs_norm) | (df_excel['tgl_input_norm'] == ""))
        ].to_dict('records')

        if not rows_excel_toko:
            rows_excel_toko = df_excel[
                (df_excel['nama_norm'] == nama_gs) &
                (df_excel['toko_norm'] == toko_gs)
            ].to_dict('records')

        if not rows_excel_toko:
            continue

        cell_address = rowcol_to_a1(r_idx + 1, target_col_idx + 1)
        
        note_lama = ""
        if r_idx < len(cell_notes) and target_col_idx < len(cell_notes[r_idx]):
            note_lama = cell_notes[r_idx][target_col_idx]

        if note_lama:
            note_baru, ada_ubah = proses_pembaharuan_teks_catatan(note_lama, rows_excel_toko)
            if ada_ubah:
                updates_notes[cell_address] = note_baru

    if updates_notes:
        print(f"--> Mengirimkan {len(updates_notes)} pembaruan Catatan (Notes) secara BATCH ke Google Sheets...")
        ws.update_notes(updates_notes)
        print(f"--> Berhasil memperbarui data pelunasan pada Google Sheets {label_proses}!")
    else:
        print("--> Tidak ada perubahan teks catatan baru yang perlu diperbarui.")

def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"--> Error file konfigurasi '{CONFIG_FILE}' tidak ditemukan!")
        return

    if not os.path.exists(CREDENTIALS_FILE):
        print(f"--> Error file '{CREDENTIALS_FILE}' tidak ditemukan!")
        return

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)

    if not config.has_section('SS'):
        print("--> Lewati section [SS] tidak ditemukan di config.conf!")
        return

    wback_status = config.get('SS', 'wback', fallback='No').strip().lower()
    if wback_status not in ['ya', 'yes', 'true', '1']:
        print("--> Lewati keterangan 'wback' pada config.conf diset 'No'. Writeback dilewati.")
        return

    print("--> Melakukan autentikasi Google API via credentials.json...")
    try:
        gc = gspread.service_account(filename=CREDENTIALS_FILE)
    except Exception as e:
        print(f"--> Error gagal melakukan autentikasi Google API: {e}")
        return

    url_irc = config.get('SS', 'url-irc', fallback='').strip()
    sn_irc = config.get('SS', 'url-irc-sn', fallback='Sheet1').strip()
    col_irc = config.get('SS', 'wback-irc-col', fallback='Nominal Nota Belum Lunas').strip()

    url_zn = config.get('SS', 'url-zn', fallback='').strip()
    sn_zn = config.get('SS', 'url-zn-sn', fallback='Form Responses 1').strip()
    col_zn = config.get('SS', 'wback-zn-col', fallback='Nominal Nota Belum Lunas').strip()

    if url_irc:
        jalankan_writeback(
            url_sheet=url_irc,
            sheet_name=sn_irc,
            col_target_name=col_irc,
            file_excel_hasil="Hasil_Ekstrak_Rincian_IRC.xlsx",
            label_proses="IRC",
            gc=gc
        )
    else:
        print("--> Lewati 'url-irc' di config.conf kosong. Writeback IRC dilewati.")

    if url_zn:
        jalankan_writeback(
            url_sheet=url_zn,
            sheet_name=sn_zn,
            col_target_name=col_zn,
            file_excel_hasil="Hasil_Ekstrak_Rincian_ZN.xlsx",
            label_proses="ZN",
            gc=gc
        )
    else:
        print("--> Lewati 'url-zn' di config.conf kosong. Writeback ZN dilewati.")

    print("\n--> Seluruh proses writeback selesai.")

if __name__ == "__main__":
    main()