import pandas as pd
import json
import requests
import os

URL = "https://1drv.ms/x/c/426713E72C49EDEB/IQDLWmSMeM5ISpQvJTI04BL7AdJu5Ck9PUYB1X4HPpBwsa8?download=1"
TEMP_FILE = "temp.xlsx"

def fetch_data():
    print("Mendownload file Excel dari OneDrive...")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    response = requests.get(URL, headers=headers)
    response.raise_for_status()
    
    with open(TEMP_FILE, "wb") as f:
        f.write(response.content)

    print("Membaca file Excel...")
    
    try:
        xls = pd.ExcelFile(TEMP_FILE)
        print(f"Berhasil membuka Excel. Daftar sheet: {xls.sheet_names}")
    except Exception as e:
        print(f"Gagal membaca format Excel: {e}")
        return

    raw_sheet_name = next((s for s in xls.sheet_names if "RAW" in s.upper() or "STOCK" in s.upper()), xls.sheet_names[0])
    soh_sheet_name = next((s for s in xls.sheet_names if "SOH" in s.upper()), xls.sheet_names[1] if len(xls.sheet_names) > 1 else xls.sheet_names[0])

    # 1. Ambil info Tanggal & Jam dari Sheet RAW
    df_raw = pd.read_excel(xls, sheet_name=raw_sheet_name, header=None)
    try:
        report_text = df_raw.iloc[0, 13] if len(df_raw.columns) > 13 else "Stock Position Report"
        date_text = df_raw.iloc[0, 10] if len(df_raw.columns) > 10 else ""
        time_text = df_raw.iloc[0, 12] if len(df_raw.columns) > 12 else ""
        report_info = f"{report_text}".strip()
        raw_date = str(date_text).strip()
        raw_time = str(time_text).strip()
    except Exception as e:
        report_info = "Stock Position Report"
        raw_date = "-"
        raw_time = "-"

    # 2. Parsing Sheet SOH Berdasarkan Kolom yang Ditemukan
    df_soh = pd.read_excel(xls, sheet_name=soh_sheet_name, header=None)
    
    # Mapping kategori sesuai struktur Excel aktual: (Col_Article, Col_Desc, Col_Qty)
    categories_map = {
        "iPhone": (2, 3, 4),          # C, D, E
        "iPad": (8, 9, 10),         # I, J, K
        "Mac": (14, 15, 16),        # O, P, Q
        "Apple Watch": (20, 21, 22),# U, V, W
        "AirPods": (26, 27, 28)     # AA, AB, AC
    }

    all_items = []
    category_summaries = {}

    start_row = 14 # Baris data mulai setelah header

    for cat, (col_art, col_desc, col_qty) in categories_map.items():
        cat_items = []
        for i in range(start_row, len(df_soh)):
            if len(df_soh.columns) <= col_qty:
                continue
                
            article = str(df_soh.iloc[i, col_art]).strip() if pd.notna(df_soh.iloc[i, col_art]) else ""
            desc = str(df_soh.iloc[i, col_desc]).strip() if pd.notna(df_soh.iloc[i, col_desc]) else ""
            qty = df_soh.iloc[i, col_qty]

            if article == "" or article.lower() == "nan" or "grand total" in article.lower():
                continue

            try:
                qty = int(qty)
            except:
                qty = 0

            item_obj = {
                "category": cat,
                "article": article,
                "description": desc,
                "qty": qty
            }
            cat_items.append(item_obj)
            all_items.append(item_obj)
        
        category_summaries[cat] = {
            "count_variants": len(cat_items),
            "total_qty": sum(item["qty"] for item in cat_items)
        }

    data = {
        "report_date": raw_date,
        "report_time": raw_time,
        "report_info": report_info,
        "summaries": category_summaries,
        "items": all_items
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)
        
    print("Berhasil memperbarui data.json")

    if os.path.exists(TEMP_FILE):
        os.remove(TEMP_FILE)

if __name__ == "__main__":
    fetch_data()
