import os
import shutil
import subprocess
import glob
import sys

dapur_dir = "Dapur"
required_dapur_files = [
                        "1_Unduh IRC ZN File.py",
                        "2_Hapus dan Filter Data.py",
                        "3_Ekstrak Komen IRC.py",
                        "3_Ekstrak Komen ZN.py",
                        "4_LookupDatas IRC.py",
                        "4_LookupDatas ZN.py",
                        "5_Cek Pelunasan IRC.py",
                        "5_Cek Pelunasan ZN.py",
                        "6_FinalisasiData IRC.py",
                        "6_FinalisasiData ZN.py",
                        "__init__.py"
                        ]

missing_items = []

if not os.path.exists(dapur_dir):
    missing_items.append(dapur_dir)
else:
    for f in required_dapur_files:
        file_path = os.path.join(dapur_dir, f)
        if not os.path.exists(file_path):
            missing_items.append(file_path)

if missing_items:
    print("--> Proses digagalkan. File atau folder berikut tidak ditemukan:")
    for item in missing_items:
        print(f"--> {item}")
    input("--> Tekan Enter untuk keluar...")
    sys.exit()

for ext in ['*.xls', '*.xlsx']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        os.remove(file)

scripts_to_run = [
                 "1_Unduh IRC ZN File.py",
                 "2_Hapus dan Filter Data.py",
                 "3_Ekstrak Komen IRC.py",
                 "3_Ekstrak Komen ZN.py",
                 "4_LookupDatas IRC.py",
                 "4_LookupDatas ZN.py",
                 "5_Cek Pelunasan IRC.py",
                 "5_Cek Pelunasan ZN.py",
                 "6_FinalisasiData IRC.py",
                 "6_FinalisasiData ZN.py"
                 ]
current_dir = os.getcwd()
os.chdir(dapur_dir)

try:
    for script in scripts_to_run:
        print(f"--> Menjalankan {script}...")
        subprocess.run([sys.executable, script], check=True)
except subprocess.CalledProcessError:
    print(f"--> Terjadi kesalahan saat menjalankan {script}. Proses dihentikan.")
    os.chdir(current_dir)
    input("--> Tekan Enter untuk keluar...")
    sys.exit()

os.chdir(current_dir)

order_files = glob.glob(os.path.join(dapur_dir, "Hasil_Ekstrak_Rincian_ZN.xlsx"))
if order_files:
    for file_path in order_files:
        file_name = os.path.basename(file_path)
        dest_path = os.path.join(current_dir, file_name)
        if os.path.exists(dest_path):
            os.remove(dest_path)
        shutil.move(file_path, dest_path)
        print(f"--> File {file_name} berhasil dipindahkan.")
else:
    print("--> Gagal: Tidak ditemukan file Hasil_Ekstrak_Rincian_ZN setelah proses.")
    
order_files = glob.glob(os.path.join(dapur_dir, "Hasil_Ekstrak_Rincian_IRC.xlsx"))
if order_files:
    for file_path in order_files:
        file_name = os.path.basename(file_path)
        dest_path = os.path.join(current_dir, file_name)
        if os.path.exists(dest_path):
            os.remove(dest_path)
        shutil.move(file_path, dest_path)
        print(f"--> File {file_name} berhasil dipindahkan.")
else:
    print("--> Gagal: Tidak ditemukan file Hasil_Ekstrak_Rincian_IRC setelah proses.")

for ext in ['*.xls', '*.xlsx']:
    for file in glob.glob(os.path.join(dapur_dir, ext)):
        os.remove(file)

print("--> Semua proses selesai dan folder Dapur telah dibersihkan.")