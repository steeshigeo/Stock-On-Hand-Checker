import os
import base64
import requests

def get_direct_download_url(share_url):
    # Konversi link publik OneDrive ke direct download URL tanpa login Azure
    base64_bytes = base64.b64encode(share_url.encode("utf-8"))
    base64_string = base64_bytes.decode("utf-8").strip("=").replace('/', '_').replace('+', '-')
    encoded_url = "u!" + base64_string
    return f"https://api.onedrive.com/v1.0/shares/{encoded_url}/root/content"

def download_excel():
    share_url = os.environ.get("ONEDRIVE_SHARE_URL", "https://1drv.ms/x/c/426713E72C49EDEB/IQDLWmSMeM5ISpQvJTI04BL7AdJu5Ck9PUYB1X4HPpBwsa8?e=4Cv2DX")
    download_url = get_direct_download_url(share_url)
    
    print("Mengunduh file Excel terbaru dari OneDrive...")
    response = requests.get(download_url, allow_redirects=True)
    
    if response.status_code == 200:
        with open("stockpositionreportdetail.xls", "wb") as f:
            f.write(response.content)
        print("Berhasil mengunduh stockpositionreportdetail.xls")
    else:
        raise Exception(f"Gagal mengunduh file, status code: {response.status_code}")

if __name__ == "__main__":
    download_excel()
