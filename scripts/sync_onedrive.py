import os
import requests

def download_excel():
    share_url = os.environ.get("ONEDRIVE_SHARE_URL", "https://1drv.ms/x/c/426713E72C49EDEB/IQDLWmSMeM5ISpQvJTI04BL7AdJu5Ck9PUYB1X4HPpBwsa8?e=4Cv2DX")
    
    sep = "&" if "?" in share_url else "?"
    download_url = f"{share_url}{sep}download=1"
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    # Buat direktori downloaded jika belum ada
    output_dir = "downloaded"
    os.makedirs(output_dir, exist_ok=True)
    file_path = os.path.join(output_dir, "stockpositionreportdetail.xls")
    
    print("Mengunduh file Excel terbaru dari OneDrive...")
    response = requests.get(download_url, headers=headers, allow_redirects=True)
    
    if response.status_code == 200:
        with open(file_path, "wb") as f:
            f.write(response.content)
        print(f"Berhasil mengunduh {file_path}")
    else:
        raise Exception(f"Gagal mengunduh file, status code: {response.status_code}")

if __name__ == "__main__":
    download_excel()
