import os
import re
import json
import base64
import requests
import pandas as pd

# Ambil dan bersihkan URL dari karakter tak terduga (petik, spasi, kurung)
raw_url = os.environ.get("ONEDRIVE_URL", "").strip()
raw_url = raw_url.strip("'\"[]()")
if raw_url and not raw_url.startswith("http://") and not raw_url.startswith("https://"):
    raw_url = "https://" + raw_url

ONEDRIVE_RAW_URL = raw_url

def download_excel_bytes(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    session = requests.Session()
    session.headers.update(headers)

    # Metode 1: Direct Parameter on Shortlink (Paling Stabil)
    try:
        direct_url = url + ("&download=1" if "?" in url else "?download=1")
        res = session.get(direct_url, timeout=30, allow_redirects=True)
        if res.status_code == 200 and len(res.content) > 500:
            print("Berhasil mengunduh Excel via Direct Shortlink Parameter.")
            return res.content
    except Exception as e:
        print(f"Metode 1 (Direct Parameter) dilewati: {e}")

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
            print("Berhasil mengunduh Excel via Redirect Resolution.")
            return res.content
    except Exception as e:
        print(f"Metode 2 (Redirect Resolution) dilewati: {e}")

    # Metode 3: Microsoft Graph API
    try:
        encoded = base64.b64encode(url.encode('utf-8')).decode('utf-8')
        url_safe = encoded.replace('+', '-').replace('/', '_').rstrip('=')
        graph_url = f"https://graph.microsoft.com/v1.0/shares/u!{url_safe}/driveItem/content"
        
        res = session.get(graph_url, timeout=30, allow_redirects=True)
        if res.status_code == 200 and len(res.content) > 500:
            print("Berhasil mengunduh Excel via Microsoft Graph API.")
            return res.content
    except Exception as e:
        print(f"Metode 3 (Graph API) dilewati: {e}")

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
        if isinstance(r, list):
            for idx, cell in enumerate(r):
                if cell and pd.notna(cell):
                    str_val = str(cell).strip()
                    if "stock position report for" in str_val.lower():
                        match = re.search(r"stock position report for\s*([0-9\-\/A-Za-z]+)", str_val, re.I)
                        if match:
                            found_date = match.group(1).strip()
                    if str_val == "Time :" and idx + 1 < len(r):
                        raw_t = str(r[idx + 1]).strip()
                        if "T" in raw_t:
                            raw_t = raw_t.split("T")[1]
                        found_time = raw_t.split(".")[0] + " WIB"

    cat_row_idx = -1
    for idx, r in enumerate(rows):
        if isinstance(r, list) and any(cell and any(k in str(cell) for k in ['iPhone', 'iPad', 'Mac']) for cell in r if pd.notna(cell)):
            cat_row_idx = idx
            break

    if cat_row_idx != -1:
        cat_row = rows[cat_row_idx]
        hdr_row_idx = cat_row_idx + 2
        categories = [(i, str(c).strip()) for i, c in enumerate(cat_row) if pd.notna(c) and str(c).strip()]

        for col_idx, cat in categories:
            items = []
            for r in range(hdr_row_idx + 1, len(rows)):
                row = rows[r]
                if len(row) > col_idx + 2:
                    art = str(row[col_idx]).strip() if pd.notna(row[col_idx]) else ""
                    desc = str(row[col_idx + 1]).strip() if pd.notna(row[col_idx + 1]) else ""
                    qty_val = row[col_idx + 2]
                    
                    if art and art != "Grand Total" and desc and pd.notna(qty_val):
                        try:
                            q = int(qty_val)
                            items.append({"article": art, "description": desc, "qty": q})
                        except ValueError:
                            pass
            if items:
                new_soh[cat] = {
                    "grand_total": sum(i["qty"] for i in items),
                    "items": items
                }
                
    return new_soh, found_date, found_time

def update_index_html(soh_data, raw_date, raw_time):
    if not os.path.exists("index.html"):
        print("File index.html tidak ditemukan!")
        return

    with open("index.html", "r", encoding="utf-8") as f:
        content = f.read()

    soh_json = json.dumps(soh_data)
    content = re.sub(r'const INITIAL_SOH_DATA = \{.*?\};', f'const INITIAL_SOH_DATA = {soh_json};', content, flags=re.DOTALL)
    
    if raw_date:
        content = re.sub(r'let dateUpdated = ".*?";', f'let dateUpdated = "{raw_date}";', content)
    if raw_time:
        content = re.sub(r'let timeUpdated = ".*?";', f'let timeUpdated = "{raw_time}";', content)

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
