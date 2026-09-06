import pandas as pd
import json
import requests
import os

URL = "https://1drv.ms/x/c/426713E72C49EDEB/IQDLWmSMeM5ISpQvJTI04BL7AdJu5Ck9PUYB1X4HPpBwsa8?download=1"
TEMP_FILE = "temp.xlsx"

def fetch_data():
    print("Mendownload file Excel dari OneDrive...")
    # Menambahkan header User-Agent agar tidak diblokir oleh sistem OneDrive
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    response = requests.get(URL, headers=headers)
    response.raise_for_status()
    
    with open(TEMP_FILE, "wb") as f:
        f.write(response.content)

    print("Membaca file Excel...")
    
    try:
        # Gunakan pd.ExcelFile untuk membaca workbook dan mendata semua sheet
        xls = pd.ExcelFile(TEMP_FILE)
        print(f"Berhasil membuka Excel. Daftar sheet yang ditemukan: {xls.sheet_names}")
    except Exception as e:
        print(f"Gagal membaca format Excel. Error: {e}")
        return

    # Cari nama sheet secara dinamis (mengabaikan typo atau spasi ekstra)
    raw_sheet_name = next((s for s in xls.sheet_names if "RAW" in s.upper() or "STOCK" in s.upper()), None)
    soh_sheet_name = next((s for s in xls.sheet_names if "SOH" in s.upper()), None)

    # Fallback jika nama sheet berubah total
    if not raw_sheet_name:
        print("Peringatan: Sheet 'RAW' tidak ditemukan. Menggunakan sheet pertama.")
        raw_sheet_name = xls.sheet_names[0]
        
    if not soh_sheet_name:
        print("Peringatan: Sheet 'SOH' tidak ditemukan. Menggunakan sheet kedua.")
        soh_sheet_name = xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0]

    print(f"=> Menggunakan sheet: '{raw_sheet_name}' untuk Data RAW")
    print(f"=> Menggunakan sheet: '{soh_sheet_name}' untuk Data SOH")

    # 1. Ambil keterangan Date & Time dari Sheet RAW
    df_raw = pd.read_excel(xls, sheet_name=raw_sheet_name, header=None)
    
    try:
        report_text = df_raw.iloc[0, 13] if len(df_raw.columns) > 13 else "Stock Position Report"
        time_text = df_raw.iloc[0, 12] if len(df_raw.columns) > 12 else ""
        report_info = f"{report_text} | Time: {time_text}".strip()
    except Exception as e:
        report_info = "Stock Position Report"
        print(f"Peringatan: Gagal mengekstrak tanggal, error: {e}")

    # 2. Parsing Sheet SOH
    df_soh = pd.read_excel(xls, sheet_name=soh_sheet_name, header=None)
    
    categories_map = {
        "iPhone": (1, 2, 3),        
        "iPad": (5, 6, 7),          
        "Mac": (9, 10, 11),         
        "Apple Watch": (13, 14, 15),
        "AirPods": (17, 18, 19)     
    }

    data = {
        "report_info": report_info,
        "categories": {}
    }

    start_row = 14

    for cat, (col_art, col_desc, col_qty) in categories_map.items():
        items = []
        for i in range(start_row, len(df_soh)):
            # Hindari error jika jumlah kolom sheet kurang dari yang diharapkan
            if len(df_soh.columns) <= col_qty:
                continue 
                
            article = str(df_soh.iloc[i, col_art]).strip() if pd.notna(df_soh.iloc[i, col_art]) else ""
            desc = str(df_soh.iloc[i, col_desc]).strip() if pd.notna(df_soh.iloc[i, col_desc]) else ""
            qty = df_soh.iloc[i, col_qty]

            if article == "" or article == "nan":
                continue

            try:
                qty = int(qty)
            except:
                qty = 0

            items.append({
                "article": article,
                "description": desc,
                "qty": qty
            })
        
        data["categories"][cat] = items

    with open("data.json", "w") as f:
        json.dump(data, f, indent=4)
        
    print("Berhasil menyimpan data ke data.json")

    if os.path.exists(TEMP_FILE):
        os.remove(TEMP_FILE)

if __name__ == "__main__":
    fetch_data()
