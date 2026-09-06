import os
import re
import json
import base64
import requests
import pandas as pd

# Ambil dan bersihkan URL OneDrive dari karakter tersembunyi
raw_url = os.environ.get("ONEDRIVE_URL", "").strip()
raw_url = raw_url.strip("'\"[]()")
if raw_url and not raw_url.startswith("http://") and not raw_url.startswith("https://"):
    raw_url = "https://" + raw_url

ONEDRIVE_RAW_URL = raw_url

def classify_product(article, description, raw_category=""):
    """
    Kategorisasi otomatis produk ke dalam kategori utama Apple & Aksesori
    """
    text = f"{article} {description}".upper()
    cat_upper = raw_category.upper()
    
    # 1. Deteksi Aksesori & Produk Non-Device
    accessory_keywords = [
        "CASE", "CHARGER", "POWER BANK", "POWERBANK", "CABLE", "ADAPTER", 
        "PROTECTOR", "FLIXMAG", "VOLTIX", "MAGXEN", "JBL", "EARPHONES", 
        "SLEEVE", "BAG", "STAND", "TEMPERED", "STRAP", "BAND", "HUB"
    ]
    if any(k in text for k in accessory_keywords) or any(k in cat_upper for k in ["ACC", "CASE", "CHARGER", "AUDIO"]):
        return "Aksesori & Lainnya"
        
    # 2. Deteksi Perangkat Utama Apple
    if "IPHONE" in text or "PHONE 1" in text:
        return "iPhone"
    elif "IPAD" in text:
        return "iPad"
    elif any(k in text for k in ["MBP", "MBA", "MACBOOK", "MAC MINI", "IMAC", "MAC STUDIO", "MAC PRO", "MBN"]):
        return "Mac"
    elif "WATCH" in text or "ULTRA" in text:
        return "Apple Watch"
    elif "AIRPODS" in text or "EARPODS" in text:
        return "AirPods"
    elif raw_category and raw_category.strip() not in ["0", "DEMO", "APPLE", "PHONE", ""]:
        return raw_category.strip()
    else:
        return "Aksesori & Lainnya"

def download_excel_bytes(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    session = requests.Session()
    session.headers.update(headers)

    # Metode 1: Direct Parameter Shortlink
    try:
        direct_url = url + ("&download=1" if "?" in url else "?download=1")
        res = session.get(direct_url, timeout=30, allow_redirects=True)
        if res.status_code == 200 and len(res.content) > 500:
            return res.content
    except Exception:
        pass

    # Metode 2: Redirect Resolution
    try:
        res = session.get(url, timeout=30, allow_redirects=True)
        final_url = res.url
        if "view.aspx" in final_url:
            download_url = final_url.replace("view.aspx", "download.aspx")
        elif "download=1" not in final_url:
            download_url = final_url + ("&download=1" if "?" in final_url else "?download=1")
        else:
            download_url = final_url

        res = session.get(download_url, timeout=30, allow_redirects=True)
        if res.status_code == 200 and len(res.content) > 500:
            return res.content
    except Exception:
        pass

    # Metode 3: Microsoft Graph API
    try:
        encoded = base64.b64encode(url.encode('utf-8')).decode('utf-8')
        url_safe = encoded.replace('+', '-').replace('/', '_').rstrip('=')
        graph_url = f"https://graph.microsoft.com/v1.0/shares/u!{url_safe}/driveItem/content"
        
        res = session.get(graph_url, timeout=30, allow_redirects=True)
        if res.status_code == 200 and len(res.content) > 500:
            return res.content
    except Exception:
        pass

    raise Exception("Gagal mengunduh file Excel dari seluruh metode OneDrive.")

def fetch_and_parse_excel(url):
    content = download_excel_bytes(url)
    
    excel_path = "temp_soh.xlsx"
    with open(excel_path, "wb") as f:
        f.write(content)
        
    xl = pd.ExcelFile(excel_path)
    all_rows = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(excel_path, sheet_name=sheet, header=None)
        all_rows.extend(df.values.tolist())
        
    if os.path.exists(excel_path):
        os.remove(excel_path)
    return all_rows

def parse_soh_data(rows):
    new_soh = {}
    found_date = None
    found_time = None
    
    for r in rows:
        if not isinstance(r, list):
            continue
            
        cells = [str(cell).strip() if pd.notna(cell) and str(cell).strip() != 'nan' else "" for cell in r]
        non_empty = [c for c in cells if c]
        if not non_empty:
            continue
            
        row_str = " ".join(non_empty)
        row_str_lower = row_str.lower()
        
        # Extract Metadata Date
        if "stock position report for" in row_str_lower:
            match = re.search(r"stock position report for\s*([0-9\-\/A-Za-z]+)", row_str, re.I)
            if match:
                found_date = match.group(1).strip()
        
        # Extract Metadata Time
        if "time" in row_str_lower:
            time_match = re.search(r'(\d{2}:\d{2}:\d{2})', row_str)
            if time_match:
                found_time = time_match.group(1) + " WIB"

        # Bypass header / metadata rows
        if any(keyword in row_str_lower for keyword in ["stock position report", "grand total", "report for", "date :", "article description", "material description"]):
            continue

        # Cari nilai Quantity dari kanan ke kiri (stok riil < 100,000 unit, membuang barcode EAN 13-digit)
        qty = None
        qty_idx = -1
        for i in range(len(cells) - 1, -1, -1):
            val = cells[i]
            if val != "" and re.match(r'^\d+$', val):
                num = int(val)
                if 0 <= num < 100000 and len(val) < 10:
                    qty = num
                    qty_idx = i
                    break
                    
        if qty is None:
            continue
            
        text_cells = [c for c in cells[:qty_idx] if c != ""]
        if not text_cells:
            continue
            
        raw_cat = text_cells[0] if len(text_cells) >= 1 else ""
        art_code = ""
        description = ""
        
        # Cari Kode Artikel SAP (misal: APPMG2U4ID/A, IBSIP2026017)
        for cell in text_cells:
            if re.search(r'[A-Z0-9]{5,}/[A-Z0-9]+', cell) or (re.search(r'^[A-Z0-9]{8,15}$', cell) and not re.match(r'^\d+$', cell)):
                art_code = cell
                break
                    
        # Cari Deskripsi Produk
        desc_candidates = [c for c in text_cells if len(c) > 5 and not re.match(r'^\d+$', c)]
        if desc_candidates:
            desc_candidates.sort(key=lambda x: len(x), reverse=True)
            description = desc_candidates[0]
            
        # Pisahkan jika format deskripsi menggabungkan Vendor & Deskripsi ("CODE / DESC")
        if " / " in description:
            parts = description.split(" / ", 1)
            if not art_code or art_code == text_cells[0]:
                art_code = parts[0].strip()
            description = parts[1].strip()
            
        if not art_code and text_cells:
            art_code = text_cells[0]
        if not description and text_cells:
            description = text_cells[-1]

        final_cat = classify_product(art_code, description, raw_cat)
        
        if final_cat not in new_soh:
            new_soh[final_cat] = {"grand_total": 0, "items": []}
            
        new_soh[final_cat]["items"].append({
            "article": art_code,
            "description": description,
            "qty": qty
        })
        new_soh[final_cat]["grand_total"] += qty
        
    return new_soh, found_date, found_time

def update_index_html(soh_data, raw_date, raw_time):
    if not os.path.exists("index.html"):
        print("File index.html tidak ditemukan!")
        return

    with open("index.html", "r", encoding="utf-8") as f:
        content = f.read()

    soh_json = json.dumps(soh_data, ensure_ascii=False)
    
    # Pengantian aman tanpa merusak blok JavaScript
    content = re.sub(
        r'const INITIAL_SOH_DATA = \{.*?\};',
        lambda m: f'const INITIAL_SOH_DATA = {soh_json};',
        content,
        flags=re.DOTALL
    )
    
    if raw_date:
        content = re.sub(
            r'let dateUpdated = ".*?";',
            lambda m: f'let dateUpdated = "{raw_date}";',
            content
        )
    if raw_time:
        content = re.sub(
            r'let timeUpdated = ".*?";',
            lambda m: f'let timeUpdated = "{raw_time}";',
            content
        )

    with open("index.html", "w", encoding="utf-8") as f:
        f.write(content)
        
    print("Berhasil memperbarui index.html dengan data SOH terbaru.")

if __name__ == "__main__":
    if not ONEDRIVE_RAW_URL:
        print("Environment variable ONEDRIVE_URL tidak diset. Melewati update auto.")
    else:
        try:
            rows = fetch_and_parse_excel(ONEDRIVE_RAW_URL)
            soh_data, raw_date, raw_time = parse_soh_data(rows)
            if soh_data and len(soh_data) > 0:
                update_index_html(soh_data, raw_date, raw_time)
            else:
                print("Data SOH hasil parsing kosong, mempertahankan data index.html yang ada.")
        except Exception as e:
            print(f"Error saat menjalankan update: {e}")
