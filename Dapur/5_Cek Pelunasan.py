import os
import re
import configparser
import pandas as pd

CONFIG_FILE = "config.conf"

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

def proses_cek_pelunasan(file_utama, ref_dict, label_proses):
    print(f"--> Cek pelunasan: {label_proses}")

    if not os.path.exists(file_utama):
        print(f"--> Lewati file utama '{file_utama}' tidak ditemukan!")
        return

    print(f"--> Membaca file utama: {file_utama}")
    df_utama = pd.read_excel(file_utama)

    kolom_bersih = [col for col in df_utama.columns if not str(col).startswith('Unnamed:')]
    df_utama = df_utama[kolom_bersih]

    cek_pelunasan_list = []

    print(f"--> Menghitung status pelunasan data {label_proses}...")
    for index, row in df_utama.iterrows():
        no_faktur_utama = bersihkan_faktur(row.get('No. Faktur', ''))
        
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

    df_utama.to_excel(file_utama, index=False)
    print(f"--> Berhasil! Kolom 'Cek Pelunasan' berhasil diperbarui pada: {file_utama}")

def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"--> Error file '{CONFIG_FILE}' tidak ditemukan!")
        return

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)

    file_arviewer = ""
    if config.has_section('DIR'):
        file_arviewer = config.get('DIR', 'arvi', fallback='').strip()
    else:
        print("--> Peringatan section [DIR] tidak ditemukan di config.conf!")

    if not file_arviewer or str(file_arviewer).strip() == "":
        print("--> Lewati kunci 'arvi' di config.conf kosong. Cek pelunasan dilewati.")
        return

    if not os.path.exists(file_arviewer):
        print(f"--> Lewati file ARVIEWER '{file_arviewer}' tidak ditemukan!")
        return

    print(f"--> Membaca Sheet 'Source' dari: {file_arviewer}")
    try:
        df_ref = pd.read_excel(file_arviewer, sheet_name="Source", skiprows=3, header=None, usecols="B:H")
    except Exception as e:
        print(f"--> Gagal membaca file ARVIEWER: {e}")
        return

    ref_dict = {}
    for idx in range(len(df_ref)):
        faktur_id = bersihkan_faktur(df_ref.iloc[idx, 0])
        
        if faktur_id and faktur_id != "nan":
            ref_dict[faktur_id] = {
                "Nilai Faktur Ref": bersihkan_angka(df_ref.iloc[idx, 5]),
                "Sisa Piutang Ref": bersihkan_angka(df_ref.iloc[idx, 6])
            }

    proses_cek_pelunasan(
        file_utama="Hasil_Ekstrak_Rincian_IRC_temp.xlsx",
        ref_dict=ref_dict,
        label_proses="IRC"
    )

    proses_cek_pelunasan(
        file_utama="Hasil_Ekstrak_Rincian_ZN_temp.xlsx",
        ref_dict=ref_dict,
        label_proses="ZN"
    )

    print("\n--> Seluruh proses pengecekan pelunasan selesai.")

if __name__ == "__main__":
    main()