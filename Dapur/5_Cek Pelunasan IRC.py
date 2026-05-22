import os
import re
import pandas as pd

FILE_UTAMA = "Hasil_Ekstrak_Rincian_IRC_temp.xlsx"
FILE_ARVIEWER = r"E:\ADM IRC AND ZN\ARVIEWER.xlsm"

def bersihkan_angka(val):
    if pd.isnull(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_clean = re.sub(r'[^\d]', '', str(val))
    return float(val_clean) if val_clean else 0.0

def bersihkan_faktur(val):
    if pd.isnull(val):
        return ""
    str_val = str(val).split('.')[0]
    return str_val.strip()

if not os.path.exists(FILE_ARVIEWER):
    print(f"--> File {FILE_ARVIEWER} tidak ditemukan!")
    exit()

print(f"--> Membaca Sheet 'Source' langsung dari ARVIEWER.xlsm...")
try:
    df_ref = pd.read_excel(FILE_ARVIEWER, sheet_name="Source", skiprows=3, header=None, usecols="B:H")
except Exception as e:
    print(f"--> Gagal membaca file ARVIEWER: {e}")
    exit()

if not os.path.exists(FILE_UTAMA):
    print(f"--> File utama {FILE_UTAMA} tidak ditemukan!")
    exit()

print(f"--> Membaca file utama: {FILE_UTAMA}")
df_utama = pd.read_excel(FILE_UTAMA)

kolom_bersih = [col for col in df_utama.columns if not str(col).startswith('Unnamed:')]
df_utama = df_utama[kolom_bersih]

ref_dict = {}
for idx in range(len(df_ref)):
    faktur_id = bersihkan_faktur(df_ref.iloc[idx, 0])
    
    if faktur_id and faktur_id != "nan":
        ref_dict[faktur_id] = {
            "Nilai Faktur Ref": bersihkan_angka(df_ref.iloc[idx, 5]),
            "Sisa Piutang Ref": bersihkan_angka(df_ref.iloc[idx, 6])
        }

cek_pelunasan_list = []

print("--> Menghitung status pelunasan data...")
for index, row in df_utama.iterrows():
    no_faktur_utama = bersihkan_faktur(row['No. Faktur'])
    
    if not no_faktur_utama or no_faktur_utama == "nan":
        cek_pelunasan_list.append("Lunas")
        continue

    if no_faktur_utama in ref_dict:
        nilai_faktur_ref = ref_dict[no_faktur_utama]["Nilai Faktur Ref"]
        sisa_piutang_ref = ref_dict[no_faktur_utama]["Sisa Piutang Ref"]
        
        nominal_terbayar = nilai_faktur_ref - sisa_piutang_ref
        
        if sisa_piutang_ref == 0:
            cek_pelunasan_list.append("Lunas")
        elif nominal_terbayar == 0:
            cek_pelunasan_list.append("Belum Dibayar")
        elif sisa_piutang_ref < nilai_faktur_ref:
            cek_pelunasan_list.append(f"Titip Bayar (Sisa Piutang: {int(sisa_piutang_ref)})")
        else:
            cek_pelunasan_list.append("Belum Dibayar")
    else:
        cek_pelunasan_list.append("Lunas")

df_utama['Cek Pelunasan'] = cek_pelunasan_list

df_utama.to_excel(FILE_UTAMA, index=False)
print(f"--> Selesai Bersih! Kolom 'Cek Pelunasan' berhasil diperbarui pada file: {FILE_UTAMA}")