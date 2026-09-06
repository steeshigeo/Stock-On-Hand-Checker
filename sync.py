import pandas as pd
import json
import requests
import os

# Link OneDrive. Parameter '?download=1' adalah trik standar agar file langsung terunduh, bukan masuk ke halaman viewer.
URL = "https://1drv.ms/x/c/426713E72C49EDEB/IQDLWmSMeM5ISpQvJTI04BL7AdJu5Ck9PUYB1X4HPpBwsa8?download=1"
TEMP_FILE = "temp.xlsx"

def fetch_data():
    print("Mendownload file Excel dari OneDrive...")
    response = requests.get(URL)
    response.raise_for_status()
    with open(TEMP_FILE, "wb") as f:
        f.write(response.content)

    print("Membaca file Excel...")
    
    # 1. Ambil keterangan Date & Time dari Sheet 'RAW StockPosition'
    df_raw = pd.read_excel(TEMP_FILE, sheet_name="RAW StockPosition", header=None)
    
    # Berdasarkan struktur CSV, Report info ada di Baris 0
    # Kolom 10: Date, Kolom 12: Time, Kolom 13: Report Text
    try:
        report_text = df_raw.iloc[0, 13] if len(df_raw.columns) > 13 else "Stock Position Report"
        time_text = df_raw.iloc[0, 12] if len(df_raw.columns) > 12 else ""
        report_info = f"{report_text} | Time: {time_text}".strip()
    except Exception as e:
        report_info = "Stock Position Report"
        print(f"Peringatan: Gagal mengekstrak tanggal, error: {e}")

    # 2. Parsing Sheet 'SOH'
    df_soh = pd.read_excel(TEMP_FILE, sheet_name="SOH", header=None)
    
    # Pemetaan kolom berdasarkan posisi (Indeks kolom Excel: A=0, B=1, dst.)
    # Format: { 'Kategori': (Col_Article, Col_Desc, Col_Qty) }
    categories_map = {
        "iPhone": (1, 2, 3),        # B, C, D
        "iPad": (5, 6, 7),          # F, G, H
        "Mac": (9, 10, 11),         # J, K, L
        "Apple Watch": (13, 14, 15),# N, O, P
        "AirPods": (17, 18, 19)     # R, S, T
    }

    data = {
        "report_info": report_info,
        "categories": {}
    }

    # Data mulai pada baris ke-15 (indeks 14) karena baris sebelumnya adalah header/Grand Total
    start_row = 14

    for cat, (col_art, col_desc, col_qty) in categories_map.items():
        items = []
        for i in range(start_row, len(df_soh)):
            article = str(df_soh.iloc[i, col_art]).strip() if pd.notna(df_soh.iloc[i, col_art]) else ""
            desc = str(df_soh.iloc[i, col_desc]).strip() if pd.notna(df_soh.iloc[i, col_desc]) else ""
            qty = df_soh.iloc[i, col_qty]

            # Lewati baris kosong
            if article == "" or article == "nan":
                continue

            # Parsing quantity
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

    # Simpan ke JSON
    with open("data.json", "w") as f:
        json.dump(data, f, indent=4)
        
    print("Berhasil menyimpan data ke data.json")

    # Hapus file temporary
    if os.path.exists(TEMP_FILE):
        os.remove(TEMP_FILE)

if __name__ == "__main__":
    fetch_data()
