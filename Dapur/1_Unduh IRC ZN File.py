import requests
from datetime import datetime

def download_spreadsheet(url, filename):
    try:
        print("--> Mengunduh file: " + filename)
        response = requests.get(url)
        response.raise_for_status()
        
        with open(filename, 'wb') as file:
            file.write(response.content)
            
        print("--> File berhasil disimpan: " + filename)
    except Exception as e:
        print("--> Terjadi kesalahan saat mengunduh: " + str(e))

if __name__ == "__main__":
    print("--> Memulai proses...")
    
    current_date = datetime.now().strftime("%d-%m-%Y")
    
    url_irc = "https://docs.google.com/spreadsheets/d/....."
    filename_irc = f"ORDER IRC JATENG {current_date}_temp.xlsx"
    
    url_zn = "https://docs.google.com/spreadsheets/d/....."
    filename_zn = f"ORDER ZN JATENG {current_date}_temp.xlsx"
    
    download_spreadsheet(url_irc, filename_irc)
    download_spreadsheet(url_zn, filename_zn)
    
    print("--> Semua proses selesai.")
