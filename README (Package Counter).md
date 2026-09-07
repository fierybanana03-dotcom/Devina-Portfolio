# Package Counter — 接件及时率 (Ketepatan Waktu Penerimaan Paket)

Notebook `PackageCounter_Cleaned.ipynb` adalah versi rapi dari notebook asli kamu.
Isinya dipecah jadi **2 bagian yang jelas**, tapi keduanya pakai **1 fungsi inti yang sama**
(`process_one_file`) supaya logika perhitungannya konsisten dan tidak dobel kode:

1. **Bagian 1 — Processing Harian**: proses **1 file** Excel (misalnya laporan 1 hari).
2. **Bagian 2 — Processing Bulk**: proses **banyak file sekaligus** dari 1 folder,
   hasilnya digabung jadi 1 tabel besar.

> Catatan: di notebook aslimu, sebenarnya hanya ada logika processing untuk 1 file
> (dites pakai `sample_file`). Belum ada bagian khusus "bulk" (loop banyak file).
> Di versi rapi ini saya **tambahkan** fungsi `process_bulk_folder()` yang memakai
> logika perhitungan yang sama persis, tinggal dijalankan ke folder berisi banyak file.

---

## Struktur Notebook

| Bagian | Isi |
|---|---|
| 1. Konfigurasi | Semua path input/output di-set di **satu tempat**, di paling atas |
| 2. Fungsi Inti | `process_one_file()` — logika perhitungan (sama persis dengan versi asli, cuma dirapikan & dikomentari) |
| 3. Processing Harian | Jalankan `process_one_file()` untuk 1 file → simpan ke Excel |
| 4. Processing Bulk | `process_bulk_folder()` — loop semua file di 1 folder, gabung, simpan ke Excel |

---

## Input

### Format file sumber (wajib)
- File Excel (`.xlsx`) hasil export dari sistem, sheet pertama.
- **Header ada di baris ke-2** (baris pertama berisi teks deskripsi filter, bukan nama kolom) — sudah ditangani otomatis lewat `header=1`.
- Kolom yang **wajib ada** (nama harus persis sama):
  - `查询日期`
  - `代理区名称`
  - `网点名称`
  - `运单号`
  - `网点卸车扫描时间`
  - `责任主体`
  - `是否及时（件）` *(tanda kurungnya full-width `（）`, bukan `()`)*

### Bagian 1 — Harian
Set 1 file di variabel:
```python
DAILY_INPUT_FILE = Path(r"...\nama_file.xlsx")
```

### Bagian 2 — Bulk
Set 1 folder (isinya banyak file `.xlsx`) di variabel:
```python
BULK_INPUT_FOLDER = Path(r"...\bulk_input")
BULK_FILE_PATTERN = "*.xlsx"   # bisa diubah, misal "接件及时率*.xlsx" kalau mau filter nama tertentu
```
Semua file yang cocok pattern di folder itu akan diproses satu-satu lalu digabung.
Kalau ada file yang gagal diproses (format beda / corrupt), file itu **di-skip** (ada pesan `[GAGAL]`) dan proses lanjut ke file berikutnya — tidak menghentikan seluruh proses.

Semua path ini ada di sel **"1. Konfigurasi"** di paling atas notebook — itu satu-satunya bagian yang perlu kamu edit setiap kali dipakai.

---

## Output

### Bagian 1 — Harian
Disimpan ke path `DAILY_OUTPUT_FILE` (default: `.../output/daily_processed.xlsx`)

### Bagian 2 — Bulk
Disimpan ke path `BULK_OUTPUT_FILE` (default: `.../output/bulk_processed_master.xlsx`)

### Kolom hasil (sama untuk kedua bagian)
Data direkap per kombinasi **Tanggal + Area Agen + Nama Outlet** (`查询日期`, `代理区名称`, `网点名称`):

| Kolom | Arti |
|---|---|
| `查询日期` | Tanggal query |
| `代理区名称` | Nama area agen |
| `网点名称` | Nama outlet/network |
| `应接件数` | Jumlah paket yang seharusnya diterima (total `运单号`) |
| `实接件数` | Jumlah paket yang benar-benar diterima (`网点卸车扫描时间` tidak kosong) |
| `留仓率` | Rasio paket yang belum diterima (`未接件数` / `应接件数`) |
| `未接件数` | Jumlah paket belum diterima (`应接件数` − `实接件数`) |
| `接件及时件数` | Jumlah paket diterima tepat waktu (`责任主体`=网点 & `是否及时（件）`=是) |
| `不及时件数` | Jumlah paket diterima tidak tepat waktu (`责任主体`=网点 & `是否及时（件）`=否) |
| `接件及时率（件）` | Rasio ketepatan waktu (`接件及时件数` / `应接件数`) |
| `网点接件及时率（件）` | Rasio ketepatan waktu versi outlet (lihat kolom di bawah) |
| `jumlah paket tepat waktu` | Jumlah paket tepat waktu, exclude `责任主体`=分拨 |
| `jumlah keseluruhan paket exclude GW, include outlet dan NULL` | Total paket exclude `责任主体`=分拨 (NULL tetap dihitung) |

---

## Cara Pakai

1. Buka `PackageCounter_Cleaned.ipynb` di Jupyter.
2. Jalankan sel pertama (install `pandas` & `openpyxl`) — cukup sekali saja kalau sudah pernah install.
3. Di sel **"1. Konfigurasi"**, ganti path sesuai kebutuhan:
   - Mau proses 1 file harian → isi `DAILY_INPUT_FILE` & `DAILY_OUTPUT_FILE`.
   - Mau proses banyak file sekaligus → isi `BULK_INPUT_FOLDER` & `BULK_OUTPUT_FILE`.
4. Jalankan sel **"2. Fungsi Inti"** (wajib, dipakai oleh kedua bagian).
5. Jalankan **Bagian 3** kalau mau proses harian, atau **Bagian 4** kalau mau proses bulk — dua-duanya independen, tidak perlu jalankan keduanya kalau cuma butuh satu.

## Requirement
- Python 3.12+
- `pandas`
- `openpyxl`

## Yang Berubah dari Notebook Asli
- Path & kode diagnostik (cek `sys.executable`, cek `site-packages`, cek isi sheet mentah) dihapus — itu cuma untuk debugging awal, tidak perlu dijalankan tiap kali pakai.
- Semua path hardcode dipindah ke satu sel **Konfigurasi** di atas, biar gampang diganti tanpa scroll cari-cari di tengah kode.
- Logika perhitungan di `process_one_file()` **tidak diubah sama sekali** — cuma ditata ulang & dikomentari.
- Ditambahkan `process_bulk_folder()` untuk memproses banyak file sekaligus, dengan penanganan error per file (file gagal di-skip, tidak menghentikan proses keseluruhan).
