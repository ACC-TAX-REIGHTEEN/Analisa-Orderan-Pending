# 🔍 Analisa Orderan Pending (Gandul)

> **Pemeriksaan otomatis status pelunasan orderan yang masih menggantung — dari cell notes Google Sheets ke update `LUNAS/` tanpa sentuhan manual**

Pipeline Python tujuh langkah yang mengunduh tracker order dari Google Sheets (IRC & ZN), mengekstrak rincian faktur dari **catatan sel (cell notes)**, mencocokkannya dengan file AR historis lokal, memeriksa status pelunasan di `ARVIEWER.xlsm`, menghasilkan laporan analisis Excel, lalu menulis kembali prefix `"LUNAS/"` ke cell notes Google Sheets secara otomatis untuk setiap faktur yang sudah terlunasi.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Format Data Historis Lokal](#-format-data-historis-lokal)
- [Konvensi Cell Note Google Sheets](#-konvensi-cell-note-google-sheets)
- [Output](#-output)
- [Setup Google Sheets API](#-setup-google-sheets-api)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

"Orderan gandul" adalah istilah untuk order yang sudah diinput di Google Sheets namun statusnya belum dikonfirmasi lunas — masih menggantung (*gandul*). Admin harus secara berkala mengecek satu per satu apakah faktur-faktur tersebut sudah terbayar.

Proyek ini mengotomasi seluruh proses itu:

```
Google Sheets order tracker
  (cell note berisi: "15 Jan 2026  15 Feb 2026  5.000.000")
                    ↓
         Ekstrak rincian faktur dari notes
                    ↓
   Lookup ke file AR historis lokal (cocokkan nominal)
                    ↓
      Cek status di ARVIEWER.xlsm (Sisa Piutang)
                    ↓
  Laporan Excel + tulis "LUNAS/" ke Google Sheets notes
```

---

## ✨ Fitur Utama

- **Parsing cell notes secara otomatis** — Mengurai isi catatan sel Google Sheets dengan regex untuk mengekstrak tanggal nota, tanggal jatuh tempo, dan nominal per rincian faktur.
- **Multi-rincian per baris** — Satu sel notes bisa berisi beberapa baris rincian; masing-masing dipecah menjadi satu baris terpisah di hasil analisis.
- **Lookup ke file historis lokal** — Mencocokkan nominal dengan file XLSX di folder bulanan (`BASE/JAN/*.xlsx`, `BASE/FEB/*.xlsx`, dst.) dalam window ±5 hari dari tanggal input.
- **Validasi dua arah** — Setelah ditemukan, faktur divalidasi: tanggal nota dan jatuh tempo dari notes harus cocok dengan data di file historis → status `BENAR` atau `SALAH`.
- **Cek pelunasan via ARVIEWER** — Membaca sheet `Source` di `ARVIEWER.xlsm` untuk menentukan status: `Lunas`, `Belum Dibayar`, atau `Titip Bayar (Sisa Piutang: X)`.
- **Kolom JT otomatis** — Menghitung umur jatuh tempo (`hari ini − Jatuh Tempo`) per baris faktur di laporan akhir.
- **Writeback `LUNAS/` ke Google Sheets** — Baris rincian dalam cell note yang fakturnya sudah lunas diberi prefix `LUNAS/` secara otomatis via batch update Google Sheets API.
- **Pemisah grup otomatis** — Laporan Excel menambahkan baris kosong sebagai pemisah visual antara kelompok order berbeda (beda sales, beda toko, atau beda tanggal input).
- **Flag writeback** — Fitur tulis-balik ke Google Sheets dapat dinonaktifkan via `wback = No` tanpa mengubah kode.
- **Multi-produk** — Memproses IRC dan ZN secara berurutan dalam satu eksekusi.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlwings requests gspread google-auth
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Skrip 4, 5, 7 | Baca Excel, filter, transformasi data |
| `openpyxl` | Skrip 3, 6 | Baca/tulis `.xlsx`, akses cell notes & comments |
| `xlwings` | Skrip 2 | Hapus baris & terapkan AutoFilter via Excel COM |
| `requests` | Skrip 1 | Unduh Google Sheets sebagai file XLSX |
| `gspread` | Skrip 7 | Baca & update cell notes Google Sheets API |
| `google-auth` | Skrip 7 | Autentikasi via Service Account |
| `configparser`, `re`, `glob`, `os`, `shutil`, `datetime` | Semua | Standard library |

### Aplikasi wajib
- **Microsoft Excel** — Wajib terinstall untuk Skrip 2 yang menggunakan `xlwings` (Excel COM) untuk operasi hapus baris dan AutoFilter. Hanya Skrip 2 yang membutuhkan Excel; skrip lainnya tidak.

> **Windows only (Skrip 2):** `xlwings` menggunakan Excel COM automation yang hanya tersedia di Windows.

---

## 📁 Struktur Folder & File

```
📦 Analisa-Orderan-Pending/
│
├── 📄 Jalankan Cek Orderan Gandul.py   ← Orkestrator utama. Jalankan ini
│
└── 📁 Dapur/                           ← Folder pipeline (jangan diubah)
    ├── 📄 __init__.py
    ├── 📄 1_Unduh File.py              ← Unduh order IRC & ZN dari Google Sheets
    ├── 📄 2_Hapus dan Filter Data.py   ← Bersihkan & filter via xlwings
    ├── 📄 3_Ekstrak Komen.py           ← Ekstrak rincian dari cell notes
    ├── 📄 4_LookupDatas.py             ← Lookup faktur ke file historis lokal
    ├── 📄 5_Cek Pelunasan.py           ← Cek status pelunasan via ARVIEWER
    ├── 📄 6_FinalisasiData.py          ← Format laporan + hitung JT
    ├── 📄 7_AddPayToSS.py              ← Writeback "LUNAS/" ke cell notes Sheets
    ├── 📄 config.conf                  ← Konfigurasi path, URL, parameter operasi
    └── 📄 credentials.json            ← Kredensial Google Service Account (rahasia!)
```

**Output yang dipindah ke folder utama:**

| File | Keterangan |
|---|---|
| `Hasil_Ekstrak_Rincian_IRC.xlsx` | Laporan analisis orderan pending IRC |
| `Hasil_Ekstrak_Rincian_ZN.xlsx` | Laporan analisis orderan pending ZN |

---

## 🚀 Cara Penggunaan

### Langkah 1 — Sesuaikan `config.conf`

Isi tiga bagian kritis:

```ini
[SS]
url-irc = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_IRC/edit
url-irc-sn = Sheet1
url-zn = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET_ZN/edit
url-zn-sn = Form Responses 1
wback = Ya

[DIR]
irc = E:\ADM IRC AND ZN\2026
zn = E:\ADM IRC AND ZN\2026
arvi = E:\ADM IRC AND ZN\ARVIEWER.xlsm

[WINGS]
irc-del = 2:19593
irc-hid = D:G
zn-del = 2:75
zn-hid = D:H
```

Lihat panduan lengkap di [Konfigurasi `config.conf`](#-konfigurasi-configconf).

### Langkah 2 — Pasang kredensial Google Sheets

Ganti isi `Dapur/credentials.json` dengan file JSON Google Service Account. Lihat [Setup Google Sheets API](#-setup-google-sheets-api).

### Langkah 3 — Jalankan

```bash
python "Jalankan Cek Orderan Gandul.py"
```

### Langkah 4 — Pantau progress

```
--> Menjalankan 1_Unduh File.py...
--> Sedang mengunduh file: ORDER IRC JATENG 23-08-2026_temp.xlsx...
--> File berhasil disimpan: ORDER IRC JATENG 23-08-2026_temp.xlsx
--> Menjalankan 2_Hapus dan Filter Data.py...
--> Memproses ORDER IRC JATENG 23-08-2026_temp.xlsx
--> Menjalankan 3_Ekstrak Komen.py...
--> Berhasil! Data IRC disimpan ke: Hasil_Ekstrak_Rincian_IRC_temp.xlsx
--> Menjalankan 4_LookupDatas.py...
--> Memulai pemindaian folder bulanan di: E:\ADM IRC AND ZN\2026
--> Menjalankan 5_Cek Pelunasan.py...
--> Membaca Sheet 'Source' dari: E:\ADM IRC AND ZN\ARVIEWER.xlsm
--> Menjalankan 6_FinalisasiData.py...
--> Menjalankan 7_AddPayToSS.py...
--> Mengirimkan 12 pembaruan Catatan secara BATCH ke Google Sheets...
--> Berhasil memperbarui data pelunasan pada Google Sheets IRC!
--> File Hasil_Ekstrak_Rincian_IRC.xlsx berhasil dipindahkan.
--> Semua proses selesai dan folder Dapur telah dibersihkan.
```

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Jalankan Cek Orderan Gandul.py]
   │
   ├─── Validasi: folder Dapur/ + 10 file syarat
   ├─── Bersihkan Dapur/ dari *.xls & *.xlsx lama
   ├─── Pindah ke direktori Dapur/
   │
   ├─── [1] 1_Unduh File.py
   │       Unduh Google Sheets IRC → ORDER IRC JATENG DD-MM-YYYY_temp.xlsx
   │       Unduh Google Sheets ZN  → ORDER ZN JATENG DD-MM-YYYY_temp.xlsx
   │       (URL diambil dari [SS] url-irc & url-zn, dengan sheet name/gid)
   │
   ├─── [2] 2_Hapus dan Filter Data.py  ← Butuh Microsoft Excel
   │       Buka setiap file via xlwings (Excel background, visible=True)
   │       IRC: hapus baris [irc-del] → sembunyikan kolom [irc-hid]
   │       ZN:  hapus baris [zn-del]  → sembunyikan kolom [zn-hid]
   │       Keduanya: set lebar kolom C = 30
   │       Terapkan AutoFilter kolom 9 (I) = blank → hanya tampilkan pending
   │       Simpan & tutup
   │
   ├─── [3] 3_Ekstrak Komen.py
   │       Buka file hasil Step 2 via openpyxl
   │       Per baris visible (tidak hidden):
   │         Baca: Tgl Input (A), Nama (B), Toko (C), Keterangan (J)
   │         Baca kolom target (H untuk IRC / K untuk ZN):
   │           - Total Kolom = baris pertama nilai sel
   │           - Gabungkan: cell comment + isi sel teks
   │         Parse regex: \d{1,2} [Nama Bulan] \d{4} \d{1,2} [Nama Bulan] \d{4} [\d\.]+
   │         → Satu baris output per rincian yang cocok
   │       → Hasil_Ekstrak_Rincian_IRC_temp.xlsx
   │       → Hasil_Ekstrak_Rincian_ZN_temp.xlsx
   │
   ├─── [4] 4_LookupDatas.py
   │       Per baris di file _temp.xlsx:
   │         Ambil bulan dari Tgl Input → cari subfolder BASE/[BULAN]/
   │         Scan file *.xlsx yang mengandung keyword "IRC"/"ZN" dalam ±5 hari
   │         Baca setiap kandidat (header=None) → cocokkan Nominal vs kolom [5]
   │         Jika cocok: ekstrak No. Faktur, Tgl Faktur, Jatuh Tempo, Nilai, Sisa
   │         Validasi: Tgl Nota == Tgl Faktur AND Tgl JT == Jatuh Tempo → BENAR/SALAH
   │       Update _temp.xlsx dengan kolom baru
   │
   ├─── [5] 5_Cek Pelunasan.py
   │       Baca ARVIEWER.xlsm sheet "Source" (skiprows=3, kolom B:H)
   │       Bangun ref_dict: {No. Faktur: {Nilai, Sisa Piutang}}
   │       Per baris _temp.xlsx:
   │         No. Faktur tidak ada di ref_dict / kosong → "Lunas"
   │         Sisa Piutang == 0 → "Lunas"
   │         Terbayar == 0 → "Belum Dibayar"
   │         0 < Sisa < Nilai → "Titip Bayar (Sisa Piutang: X)"
   │       Tambah kolom "Cek Pelunasan" → update _temp.xlsx
   │
   ├─── [6] 6_FinalisasiData.py
   │       Baca _temp.xlsx → buat workbook baru bersih
   │       Sisipkan baris kosong sebagai pemisah antar kelompok order
   │       Hitung kolom JT = (hari ini − Jatuh Tempo).days + " Hari"
   │       Auto-fit lebar semua kolom
   │       → Hasil_Ekstrak_Rincian_IRC.xlsx
   │       → Hasil_Ekstrak_Rincian_ZN.xlsx
   │
   ├─── [7] 7_AddPayToSS.py  (skip jika wback = No)
   │       Baca file .xlsx hasil Step 6
   │       Autentikasi Google API via credentials.json
   │       Buka Google Sheets IRC & ZN
   │       Ambil semua cell notes via metadata API (batch, efisien)
   │       Per baris Google Sheets (Nama + Toko + Tgl cocok):
   │         Parse setiap baris di cell note
   │         Cocokkan nominal note ke nominal di Excel
   │         Jika cocok & Cek Pelunasan == "Lunas" → tambah prefix "LUNAS/"
   │       Kirim perubahan notes secara BATCH ke Google Sheets
   │
   ├─── Pindah kembali ke direktori utama
   ├─── Pindahkan Hasil_Ekstrak_Rincian_IRC.xlsx → folder utama
   ├─── Pindahkan Hasil_Ekstrak_Rincian_ZN.xlsx → folder utama
   └─── Bersihkan Dapur/ dari semua *.xls & *.xlsx
```

---

## 🔍 Detail Tiap Skrip

### Skrip 1 — `1_Unduh File.py`

Mengunduh dua Google Sheets ke folder `Dapur/`. URL dikonversi otomatis ke export endpoint:

```
https://docs.google.com/spreadsheets/d/{ID}/export?format=xlsx&sheet={NAMA_SHEET}
```

Nama file output menyertakan tanggal hari ini: `ORDER IRC JATENG 23-08-2026_temp.xlsx`.

---

### Skrip 2 — `2_Hapus dan Filter Data.py`

Mengoperasikan file XLSX yang diunduh via **xlwings** (Excel COM). Tiga operasi utama:

1. **Hapus baris** (`irc-del` / `zn-del`) — Menghapus baris dalam range yang ditentukan untuk membuang data header Google Forms atau baris historis yang tidak relevan.
2. **Sembunyikan kolom** (`irc-hid` / `zn-hid`) — Menyembunyikan kolom-kolom form metadata yang tidak diperlukan untuk analisis.
3. **AutoFilter kolom I = blank** — Hanya menampilkan baris yang kolom I-nya kosong, yaitu order yang belum memiliki status konfirmasi (masih pending/gandul).

> ⚠️ Skrip ini membuka Excel secara **visible** (`visible=True`) — jendela Excel akan muncul saat proses berlangsung dan menutup sendiri setelah selesai.

---

### Skrip 3 — `3_Ekstrak Komen.py`

Membaca file hasil Step 2 menggunakan `openpyxl` (bukan Excel). Memproses setiap baris yang **tidak disembunyikan** oleh AutoFilter.

**Sumber data per baris:**

| Kolom | Indeks | Field |
|---|---|---|
| A | 0 | Tgl Input (diformat ulang ke `DD MON YY`) |
| B | 1 | Nama (sales) |
| C | 2 | Toko (nama toko pelanggan) |
| H | 7 | Kolom target IRC (cell value + cell comment) |
| K | 10 | Kolom target ZN (cell value + cell comment) |
| J | 9 | Keterangan |

**Logika ekstraksi rincian dari cell note:**

Isi cell note digabungkan dengan isi cell value (jika berupa teks), lalu diparsing dengan regex:

```
Pattern: (\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+(\d{1,2}\s+[A-Za-z]+\s+\d{4})\s+([\d\.]+)
                 Tgl Nota                         Tgl Jatuh Tempo          Nominal
```

Contoh isi note yang valid:
```
15 Januari 2026 15 Februari 2026 5.000.000
01 Feb 2026 01 Mar 2026 3.500.000
```

Jika tidak ada match → baris tetap dimasukkan dengan `Status = "Tidak Ada Rincian"`.
Jika match ditemukan → satu baris output per match (satu nota per baris).

---

### Skrip 4 — `4_LookupDatas.py`

Mencocokkan setiap rincian faktur ke file historis lokal. Strategi pencarian:

```
Tgl Input: "15 AGU 26"
  → Bulan: AGU
  → Cari folder: BASE/AGU/
  → Filter file: harus mengandung "IRC"/"ZN" di nama
  → Window tanggal: 10 AGU 26 s.d. 20 AGU 26 (±5 hari)
  → Per file kandidat: baca semua baris (header=None)
    → Cari baris dengan kolom[5] == Nominal Rincian
    → Jika ketemu: ambil kolom[0,1,3,5,6] sebagai data faktur
```

**Validasi setelah ditemukan:**

| Kondisi | Status |
|---|---|
| Tgl Nota == Tgl Faktur AND Tgl JT == Jatuh Tempo | `BENAR` |
| Salah satu tidak cocok | `SALAH` |
| Nominal tidak ditemukan di manapun | `SALAH` (data match kosong) |

---

### Skrip 5 — `5_Cek Pelunasan.py`

Membaca sheet `Source` di `ARVIEWER.xlsm` (skip 3 baris, kolom B hingga H, tanpa header) untuk membangun kamus referensi pelunasan:

```python
ref_dict = {
    "100001": {"Nilai Faktur Ref": 5000000.0, "Sisa Piutang Ref": 0.0},
    "100002": {"Nilai Faktur Ref": 3500000.0, "Sisa Piutang Ref": 1500000.0},
    ...
}
```

**Logika penentuan status:**

| Kondisi | Status `Cek Pelunasan` |
|---|---|
| No. Faktur tidak ada di ARVIEWER (sudah hilang dari sistem AR) | `Lunas` |
| No. Faktur kosong / NaN | `Lunas` |
| `Sisa Piutang == 0` | `Lunas` |
| `Nilai Faktur − Sisa Piutang == 0` (belum ada bayaran) | `Belum Dibayar` |
| `0 < Sisa Piutang < Nilai Faktur` | `Titip Bayar (Sisa Piutang: X)` |

---

### Skrip 6 — `6_FinalisasiData.py`

Membuat workbook baru bersih dari file `_temp.xlsx`:

- **Pemisah grup** — Baris kosong disisipkan setiap kali kombinasi `Tgl Input + Nama + Toko` berubah, memudahkan pembacaan visual.
- **Kolom JT** — Dihitung sebagai `(datetime.now() − Jatuh Tempo).days` → ditampilkan sebagai `"45 Hari"`. Jika tanggal tidak bisa diparsing → dikosongkan.
- **Auto-fit** — Setiap kolom disesuaikan lebarnya dengan konten terpanjang + 3 karakter margin, minimum 10.

---

### Skrip 7 — `7_AddPayToSS.py`

Menulis kembali informasi pelunasan ke cell notes Google Sheets. Hanya dijalankan jika `wback = Ya`.

**Pencocokan baris Google Sheets ke Excel:**
- Primary: `Nama + Toko + Tgl Input` (setelah normalisasi)
- Fallback: `Nama + Toko` saja (jika tanggal tidak cocok)

**Logika pembaruan per baris note:**
```
Untuk setiap baris dalam cell note:
  1. Cek apakah sudah ada prefix "LUNAS/" → jika ya, lewati
  2. Parse baris dengan regex (Tgl Nota, Tgl JT, Nominal)
  3. Cocokkan nominal ke data Excel (toleransi < 0.01)
  4. Jika cocok dan Cek Pelunasan == "Lunas":
       → prefix baris dengan "LUNAS/ "
       → tandai ada perubahan
```

**Pengambilan notes via metadata API:**
Notes diambil secara batch menggunakan `spreadsheet.fetch_sheet_metadata()` dengan fields `sheets(data(rowData(values(note))))` — lebih efisien daripada mengambil note per sel.

**Pengiriman perubahan:**
Hanya sel yang benar-benar berubah yang dikirim, dalam satu panggilan `ws.update_notes(dict)`.

---

## ⚙️ Konfigurasi `config.conf`

### `[SS]` — Google Sheets order tracker

```ini
[SS]
url-irc = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
url-irc-sn = Sheet1
url-zn = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
url-zn-sn = Form Responses 1
wback = Ya
```

| Key | Keterangan |
|---|---|
| `url-irc` | URL Google Sheets tracker order IRC. Kosong → seluruh proses IRC di-skip |
| `url-irc-sn` | Nama sheet atau GID angka (contoh: `Sheet1` atau `123456789`) |
| `url-zn` | URL Google Sheets tracker order ZN |
| `url-zn-sn` | Nama sheet ZN |
| `wback` | `Ya` → aktifkan writeback ke Google Sheets; `No` → hanya hasilkan laporan Excel |

**Kunci opsional untuk writeback:**

| Key | Default | Keterangan |
|---|---|---|
| `wback-irc-col` | `Nominal Nota Belum Lunas` | Nama kolom target di Google Sheets IRC |
| `wback-zn-col` | `Nominal Nota Belum Lunas` | Nama kolom target di Google Sheets ZN |

---

### `[DIR]` — Path lokal

```ini
[DIR]
irc = E:\ADM IRC AND ZN\2026
zn = E:\ADM IRC AND ZN\2026
arvi = E:\ADM IRC AND ZN\ARVIEWER.xlsm
```

| Key | Keterangan |
|---|---|
| `irc` | Folder tahun yang berisi subfolder bulan dengan file AR historis IRC |
| `zn` | Folder tahun yang berisi subfolder bulan dengan file AR historis ZN |
| `arvi` | Path absolut ke `ARVIEWER.xlsm` untuk pengecekan saldo piutang |

---

### `[WINGS]` — Parameter operasi xlwings

```ini
[WINGS]
irc-del = 2:19593
irc-hid = D:G
zn-del = 2:75
zn-hid = D:H
```

| Key | Keterangan |
|---|---|
| `irc-del` | Range baris yang dihapus dari file IRC (format: `baris_awal:baris_akhir`) |
| `irc-hid` | Kolom yang disembunyikan di file IRC (format kolom Excel: `D:G`) |
| `zn-del` | Range baris yang dihapus dari file ZN |
| `zn-hid` | Kolom yang disembunyikan di file ZN |

> **Panduan pengaturan `xxx-del`:** Sesuaikan range dengan jumlah baris header/metadata di Google Sheets masing-masing. Gunakan angka besar (seperti `2:19593`) untuk memastikan semua baris non-data ikut terhapus. Cek terlebih dahulu berapa baris di spreadsheet sebelum menetapkan nilai ini.

---

## 📂 Format Data Historis Lokal

Skrip 4 mencari faktur di struktur folder berikut:

```
[DIR] irc atau zn
└── [BULAN]/            ← contoh: JAN, FEB, MAR, ..., DES
    ├── ...IRC 15 JAN 26.xlsx
    ├── ...IRC 16 JAN 26.xlsx
    └── ...ZN 15 JAN 26.xlsx
```

**Syarat nama file:**
- Harus mengandung keyword `IRC` atau `ZN` (case-sensitive)
- Harus mengandung tanggal dalam format `\d{1,2}\s+[A-Za-z]+\s+\d{2}` (mis. `15 JAN 26`)
- File yang tanggalnya berada di luar window ±5 hari dari `Tgl Input` order akan diabaikan

**Syarat struktur isi file (tanpa header):**

| Indeks kolom | Isi |
|---|---|
| 0 | No. Faktur |
| 1 | Tgl Faktur |
| 3 | Jatuh Tempo |
| 5 | Nilai Faktur ← digunakan untuk pencocokan nominal |
| 6 | Sisa Piutang |

---

## 📝 Konvensi Cell Note Google Sheets

Cell notes di kolom target Google Sheets (biasanya `Nominal Nota Belum Lunas`) harus mengikuti format berikut agar dapat diparsing:

### Format satu rincian

```
DD Bulan YYYY  DD Bulan YYYY  Nominal
```

Contoh:
```
15 Januari 2026 15 Februari 2026 5.000.000
```

### Format multi-rincian (satu baris per faktur)

```
15 Januari 2026 15 Februari 2026 5.000.000
01 Februari 2026 01 Maret 2026 3.500.000
10 Februari 2026 10 Maret 2026 2.000.000
```

### Setelah writeback (faktur lunas)

```
LUNAS/ 15 Januari 2026 15 Februari 2026 5.000.000
01 Februari 2026 01 Maret 2026 3.500.000
```

> Baris yang sudah memiliki prefix `LUNAS/` tidak akan diproses ulang di run berikutnya.

---

## 📤 Output

### `Hasil_Ekstrak_Rincian_IRC.xlsx` & `Hasil_Ekstrak_Rincian_ZN.xlsx`

Laporan analisis orderan pending, satu baris per rincian faktur:

| Kolom | Keterangan |
|---|---|
| `Tgl Input` | Tanggal order diinput (`DD MON YY`) |
| `Nama` | Nama sales |
| `Toko` | Nama toko/pelanggan |
| `Total Kolom K` | Total nominal dari nilai sel (baris pertama) |
| `Tgl Nota` | Tanggal nota dari cell note |
| `Tgl Jth Tempo` | Tanggal jatuh tempo dari cell note |
| `Nominal Rincian` | Nominal per rincian faktur |
| `Status` | `OK`, `Check Format Angka`, atau `Tidak Ada Rincian` |
| `Keterangan` | Isi kolom Keterangan dari order |
| `No. Faktur` | Hasil lookup ke file historis |
| `Tgl Faktur` | Tanggal faktur dari file historis |
| `Jatuh Tempo` | Tanggal JT dari file historis |
| `Nilai Faktur` | Nilai faktur dari file historis |
| `Sisa Piutang` | Sisa piutang dari file historis |
| `Status Validasi` | `BENAR` (tanggal cocok) atau `SALAH` |
| `Cek Pelunasan` | `Lunas` / `Belum Dibayar` / `Titip Bayar (Sisa: X)` |
| `JT` | Umur jatuh tempo: `"45 Hari"` |

Baris kosong disisipkan sebagai pemisah antar kelompok order yang berbeda.

---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → buat/pilih project.
2. Aktifkan **Google Sheets API** dan **Google Drive API**.
3. Buka **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Di tab **Keys** → buat key baru tipe **JSON** → file terunduh otomatis.

### 2. Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON yang diunduh.

### 3. Berikan akses ke Google Sheets

Tambahkan `client_email` dari `credentials.json` sebagai **Editor** di:
- Google Sheets tracker IRC (`url-irc`)
- Google Sheets tracker ZN (`url-zn`)

---

## 🛠️ Troubleshooting

### ❌ `File atau folder berikut tidak ditemukan: Dapur/1_Unduh File.py`
Pastikan seluruh isi folder `Dapur/` ada dan lengkap (10 file syarat).

### ❌ `Terjadi kesalahan saat menjalankan 2_Hapus dan Filter Data.py`
Skrip 2 membutuhkan Microsoft Excel terinstall. Pastikan Excel bisa dibuka secara normal. Jika Excel sedang terbuka dengan file lain, tutup dulu sebelum menjalankan pipeline.

### ❌ `Lewati file Excel dengan pola 'ORDER IRC JATENG*.xlsx' tidak ditemukan`
Step 1 gagal mengunduh file. Cek: (1) `url-irc` di `config.conf` sudah diisi dengan benar; (2) Google Sheets dapat diakses publik (atau service account memiliki akses); (3) koneksi internet aktif.

### ❌ `Hasil_Ekstrak_Rincian_IRC_temp.xlsx` kosong (0 baris data)
Tidak ada baris visible setelah Step 2. Kemungkinan: (1) AutoFilter kolom I tidak menemukan baris pending; (2) range `irc-del` menghapus terlalu banyak baris. Periksa isi Google Sheets dan sesuaikan `irc-del`.

### ❌ Semua baris `Status Validasi = SALAH`
Folder historis tidak ditemukan atau format nama file tidak mengandung keyword `IRC`/`ZN` dan tanggal yang dapat diparsing. Cek path di `[DIR] irc`/`[DIR] zn` dan struktur subfolder bulan.

### ❌ `Lewati folder direktori '...' tidak ditemukan di sistem!`
Path `[DIR] irc` atau `[DIR] zn` tidak valid. Gunakan path absolut dan pastikan folder tersebut ada.

### ❌ `Lewati file ARVIEWER '...' tidak ditemukan!`
Path `[DIR] arvi` tidak valid atau file tidak ada. Pastikan pipeline ARVIEWER sudah dijalankan dan menghasilkan `ARVIEWER.xlsm` di path yang dikonfigurasi.

### ❌ `Lewati keterangan 'wback' pada config.conf diset 'No'. Writeback dilewati`
Ini bukan error — ini berarti fitur writeback sengaja dinonaktifkan. Ubah ke `wback = Ya` untuk mengaktifkan.

### ❌ `Error gagal melakukan autentikasi Google API`
`credentials.json` tidak valid atau tidak ada. Pastikan file berisi JSON Service Account yang lengkap dengan `private_key`.

### ❌ Tidak ada perubahan notes di Google Sheets setelah writeback
Kemungkinan: (1) Tidak ada baris yang `Cek Pelunasan == "Lunas"`; (2) Pencocokan `Nama + Toko + Tgl Input` gagal — periksa apakah format tanggal di Google Sheets cocok setelah normalisasi; (3) Nominal di notes tidak cocok dengan nominal di Excel (selisih > 0.01).

---

## 📌 Catatan Penting

- **Excel harus bisa berjalan di background** — Step 2 membuka Excel secara otomatis (`visible=True`). Jangan tutup jendela Excel yang muncul secara manual; biarkan skrip menutupnya sendiri.
- **Cell note TIDAK pernah dihapus** — Writeback hanya menambah prefix `LUNAS/` di depan baris. Baris yang belum lunas tidak disentuh. Ini memastikan data historis tetap terjaga.
- **Sekali lunas, selalu lunas** — Baris note dengan `LUNAS/` di run sebelumnya akan dibiarkan oleh run berikutnya, meski faktur hilang dari ARVIEWER.
- **`credentials.json` bersifat rahasia** — Tambahkan ke `.gitignore`. Jangan commit ke repositori publik.
- **Pipeline membutuhkan data ARVIEWER terkini** — Jalankan pipeline ARVIEWER terlebih dahulu untuk memastikan `ARVIEWER.xlsm` memiliki data saldo piutang yang terbaru.
- **Format cell note harus konsisten** — Regex hanya mengenali format `DD [Nama Bulan] YYYY DD [Nama Bulan] YYYY [Nominal]`. Format yang berbeda (misalnya `DD/MM/YYYY`) tidak akan terdeteksi.
- **Output files ditimpa setiap run** — `Hasil_Ekstrak_Rincian_IRC.xlsx` dan `Hasil_Ekstrak_Rincian_ZN.xlsx` di folder utama akan diganti setiap kali pipeline dijalankan.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
