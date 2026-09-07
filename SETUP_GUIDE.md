# Cek Stock Apple — Auto-Sync dari OneDrive

Alur: **Excel (OneDrive) → GitHub Actions (tiap ±2 menit) → Netlify/Vercel**

Sama seperti alur di presentasi referensimu (Excel → Google Drive → GitHub Actions → Netlify), tapi
sumber datanya OneDrive dan pakai Microsoft Graph API sebagai ganti Google Drive API.

Tim toko cukup edit `stockpositionreportdetail.xls` di OneDrive seperti biasa. Sistem lain bekerja
otomatis di belakang layar.

---

## Struktur Project

```
├── .github/workflows/sync-onedrive.yml   # penjadwal otomatis (tiap 2 menit)
├── scripts/sync_onedrive.py              # unduh file terbaru dari OneDrive
├── scripts/build_static.py               # olah Excel → data.json
├── requirements.txt                      # dependency Python
├── index.html                            # website SOH Local
└── data.json                             # data hasil olahan (dibuat otomatis, root repo — sejajar index.html)
```

> File ini disusun supaya sejajar dengan `index.html` yang sudah ada di repo kamu — tidak perlu
> folder `public/` atau ubah setting Vercel/Netlify apa pun.

---

## Step 1 — Buat Repository GitHub

1. Buat repository baru, misal `digimap-tp3-stock`.
2. Upload semua file di paket ini ke branch `main`.
3. **Jangan** upload credential Microsoft ke repository. Semua kredensial disimpan sebagai
   **GitHub Secrets** (lihat Step 3).

---

## Step 2 — Daftarkan Aplikasi di Azure (identitas "robot")

Ini setara dengan Service Account di Google — sebuah identitas bot yang boleh membaca file Excel-mu
tanpa perlu login manual tiap kali.

1. Buka [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**
   → **New registration**.
2. Beri nama, misal `Digimap Stock Sync Bot`. Klik **Register**.
3. Catat dua nilai dari halaman **Overview**:
   - **Application (client) ID**
   - **Directory (tenant) ID**
4. Buat secret: menu **Certificates & secrets** → **New client secret** → catat **Value**-nya
   (hanya muncul sekali, langsung disalin).
5. Beri izin baca file: menu **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Application permissions** → cari `Files.Read.All` → tambahkan.
6. Klik **Grant admin consent** (perlu akses admin tenant `mapcoid` — kalau kamu bukan admin,
   minta tim IT untuk klik tombol ini).

> Kalau tim IT tidak mengizinkan pembuatan App Registration atau admin consent, opsi ini buntu —
> tanyakan dulu ke IT sebelum lanjut, supaya tidak kerja dua kali.

---

## Step 3 — Simpan Kredensial sebagai GitHub Secrets

Di repository GitHub → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**,
tambahkan 4 secret berikut:

| Nama Secret | Isi |
|---|---|
| `MS_TENANT_ID` | Directory (tenant) ID dari Step 2 |
| `MS_CLIENT_ID` | Application (client) ID dari Step 2 |
| `MS_CLIENT_SECRET` | Client secret Value dari Step 2 |
| `ONEDRIVE_SHARE_URL` | Link share file `stockpositionreportdetail.xls` (yang sudah kamu tes bisa dibuka tanpa login) |

---

## Step 4 — Uji Coba Workflow

1. Buka tab **Actions** di repository → pilih workflow **Sync Stock from OneDrive**.
2. Klik **Run workflow** untuk memaksa jalan sekali secara manual.
3. Kalau semua langkah hijau (Success), berarti koneksi ke OneDrive dan proses build sudah berhasil,
   dan `public/data.json` otomatis ter-update + ter-commit.
4. Kalau merah (Failed), buka detail log — biasanya karena: link belum "Anyone with the link", izin
   `Files.Read.All` belum di-admin-consent, atau salah salin Tenant/Client ID.

---

## Step 5 — Deploy

Kamu sudah punya project Vercel yang terhubung ke repo ini — **tidak perlu setting apa pun lagi**.
Begitu GitHub Actions commit perubahan `data.json` ke branch `main`, Vercel otomatis mendeteksi
commit baru dan publish ulang website-nya sendiri (ini perilaku default Vercel untuk repo yang
sudah terhubung, tidak butuh konfigurasi tambahan karena `index.html` dan `data.json` sama-sama
ada di root repo).

---

## Alur Operasional Harian (yang dilakukan tim toko)

1. **Edit** — ubah stok/harga di file `stockpositionreportdetail.xls` seperti biasa.
2. **Save** — simpan di lokasi OneDrive yang sama (jangan hapus lalu upload file baru).
3. **Tunggu ±2 menit** — GitHub Actions otomatis mengecek dan mengambil versi terbaru.
4. **Website update otomatis** — Netlify/Vercel publish ulang begitu ada perubahan data.

Staf tidak perlu buka GitHub atau edit website sama sekali.

---

## Yang Perlu Diperhatikan

- **Bukan real-time instan.** ±2 menit adalah target interval GitHub Actions, bukan jaminan
  detik-per-detik — persis seperti catatan di presentasi referensimu untuk sistem 5 menitnya.
- **Jangan hapus file lalu upload file baru dengan nama sama** kalau strukturnya beda — sistem
  mengacu pada link share yang sama; kalau link berubah, `ONEDRIVE_SHARE_URL` di GitHub Secrets
  perlu diperbarui.
- **Kredensial aman**: Client Secret hanya tersimpan di GitHub Secrets, tidak pernah ada di kode.
- **Bisa ditelusuri**: setiap sync tercatat di tab Actions dan riwayat commit repository — kalau ada
  masalah, tinggal cek log di situ.
- Kalau butuh update paksa segera (tidak mau nunggu 2 menit), buka tab **Actions** → **Run workflow**.

---

## Kalau Nanti Butuh Instant Sync Sungguhan (bukan polling)

Pendekatan di atas adalah polling interval (sama seperti presentasi referensimu, hanya jauh lebih
sering: 2 menit vs 5 menit). Kalau suatu saat butuh update yang benar-benar terpicu instan begitu file
berubah (bukan menunggu interval), itu perlu Microsoft Graph **webhook subscription** + endpoint
penerima notifikasi (misal Netlify Function) yang memicu `repository_dispatch` ke GitHub Actions.
Ini jauh lebih kompleks (perlu App Registration tambahan, endpoint publik, dan renewal subscription
berkala) — beri tahu kalau nanti mau upgrade ke arah ini.
