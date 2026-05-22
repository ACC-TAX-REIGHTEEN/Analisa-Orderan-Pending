import os
import glob
import re
import pandas as pd
from datetime import datetime, timedelta

FILE_UTAMA = "Hasil_Ekstrak_Rincian_ZN_temp.xlsx"
FOLDER_BASE = r"E:\ADM IRC AND ZN\2026"

BULAN_MAP = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MEI": 5, "JUN": 6,
    "JUL": 7, "AGU": 8, "SEP": 9, "OKT": 10, "NOV": 11, "DES": 12
}
BULAN_REV = {v: k for k, v in BULAN_MAP.items()}

BULAN_OUTPUT = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "Mei", 6: "Jun",
    7: "Jul", 8: "Agu", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"
}

def parse_custom_date(date_str):

    try:
        parts = str(date_str).strip().split()
        if len(parts) != 3:
            return None
        day = int(parts[0])
        month = BULAN_MAP[parts[1].upper()]
        year = int("20" + parts[2])
        return datetime(year, month, day)
    except:
        return None

def format_ke_tanggal_standar(val_tgl):
    if pd.isnull(val_tgl) or str(val_tgl).strip() == "":
        return ""
        
    if isinstance(val_tgl, datetime):
        day = f"{val_tgl.day:02d}"
        month_str = BULAN_OUTPUT[val_tgl.month]
        return f"{day} {month_str} {val_tgl.year}"
        
    tgl_str = str(val_tgl).strip()
    if "-" in tgl_str:
        try:
            tgl_clean = tgl_str.split()[0]
            dt = datetime.strptime(tgl_clean, "%Y-%m-%d")
            day = f"{dt.day:02d}"
            month_str = BULAN_OUTPUT[dt.month]
            return f"{day} {month_str} {dt.year}"
        except:
            pass
            
    return tgl_str

def bersihkan_angka(val):
    if pd.isnull(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_clean = re.sub(r'[^\d]', '', str(val))
    return float(val_clean) if val_clean else 0.0

def ambil_angka_tanggal_dari_teks(teks):
    match = re.search(r'\d{1,2}\s+(?:JAN|FEB|MAR|APR|MEI|JUN|JUL|AGU|SEP|OKT|NOV|DES)\s+\d{2}', teks, re.IGNORECASE)
    if match:
        return match.group(0)
    return None

if not os.path.exists(FILE_UTAMA):
    print(f"--> File utama {FILE_UTAMA} tidak ditemukan!")
    exit()

print(f"--> Membaca file utama: {FILE_UTAMA}")
df_utama = pd.read_excel(FILE_UTAMA)

no_faktur_list = []
tgl_faktur_list = []
jatuh_tempo_list = []
nilai_faktur_list = []
sisa_piutang_list = []
status_list = []

print("--> Memulai pemindaian folder bulanan...")

for index, row in df_utama.iterrows():
    tgl_input_str = str(row['Tgl Input']).strip()
    tgl_nota_raw = row['Tgl Nota']
    tgl_jth_tempo_raw = row['Tgl Jth Tempo']
    nominal_rincian = bersihkan_angka(row['Nominal Rincian'])

    dt_input = parse_custom_date(tgl_input_str)
    
    found = False
    data_match = {
        "No. Faktur": "", "Tgl Faktur": "", "Jatuh Tempo": "", 
        "Nilai Faktur": "", "Sisa Piutang": "", "Status": "SALAH"
    }

    parts_tgl = tgl_input_str.split()
    
    if dt_input and len(parts_tgl) == 3 and nominal_rincian > 0:
        bulan_folder_nama = parts_tgl[1].upper()
        path_folder_bulan = os.path.join(FOLDER_BASE, bulan_folder_nama)
        
        if os.path.exists(path_folder_bulan):
            semua_file_di_bulan = glob.glob(os.path.join(path_folder_bulan, "*.xlsx"))
            files_kandidat = []
            
            tgl_awal_toleransi = dt_input - timedelta(days=5)
            tgl_akhir_toleransi = dt_input + timedelta(days=5)
            
            for f_path in semua_file_di_bulan:
                nama_file = os.path.basename(f_path)
                
                if "ZN" in nama_file:
                    tgl_file_str = ambil_angka_tanggal_dari_teks(nama_file)
                    if tgl_file_str:
                        dt_file = parse_custom_date(tgl_file_str)
                        if dt_file and (tgl_awal_toleransi <= dt_file <= tgl_akhir_toleransi):
                            files_kandidat.append(f_path)
            
            for f_ref in files_kandidat:
                try:
                    df_ref = pd.read_excel(f_ref, header=None)
                    
                    for _, row_ref in df_ref.iterrows():
                        if len(row_ref) < 7: 
                            continue
                            
                        ref_nilai = bersihkan_angka(row_ref[5])
                        
                        if ref_nilai == nominal_rincian:
                            tgl_faktur_standar = format_ke_tanggal_standar(row_ref[1])
                            jatuh_tempo_standar = format_ke_tanggal_standar(row_ref[3])
                            
                            data_match["No. Faktur"] = row_ref[0]
                            data_match["Tgl Faktur"] = tgl_faktur_standar
                            data_match["Jatuh Tempo"] = jatuh_tempo_standar
                            data_match["Nilai Faktur"] = row_ref[5]
                            data_match["Sisa Piutang"] = row_ref[6]
                            
                            nota_clean = str(tgl_nota_raw).strip().lower()
                            faktur_clean = str(tgl_faktur_standar).strip().lower()
                            
                            jth_tempo_clean = str(tgl_jth_tempo_raw).strip().lower()
                            ref_jth_tempo_clean = str(jatuh_tempo_standar).strip().lower()
                            
                            if nota_clean == faktur_clean and jth_tempo_clean == ref_jth_tempo_clean:
                                data_match["Status"] = "BENAR"
                            else:
                                data_match["Status"] = "SALAH"
                                
                            found = True
                            break
                    if found: break
                except:
                    continue

    no_faktur_list.append(data_match["No. Faktur"])
    tgl_faktur_list.append(data_match["Tgl Faktur"])
    jatuh_tempo_list.append(data_match["Jatuh Tempo"])
    nilai_faktur_list.append(data_match["Nilai Faktur"])
    sisa_piutang_list.append(data_match["Sisa Piutang"])
    status_list.append(data_match["Status"])

df_utama['No. Faktur'] = no_faktur_list
df_utama['Tgl Faktur'] = tgl_faktur_list
df_utama['Jatuh Tempo'] = jatuh_tempo_list
df_utama['Nilai Faktur'] = nilai_faktur_list
df_utama['Sisa Piutang'] = sisa_piutang_list
df_utama['Status Validasi'] = status_list

df_utama.to_excel(FILE_UTAMA, index=False)
print(f"--> Berhasil! File '{FILE_UTAMA}' telah diperbarui dengan data XLOOKUP berformat standar.")