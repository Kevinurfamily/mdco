import io
import os
import time
import logging
import requests
import pandas as pd
from pathlib import Path
from io import BytesIO, StringIO

# =============================
# KONFIGURASI DIREKTORI & TARGET
# =============================
OUTPUT_DIR = Path("result_copilot_kb")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Path sertifikat SSL lokal (jika diperlukan untuk jaringan kantor)
CA_BUNDLE_PATH = r"C:\Users\90006790\OneDrive - PT. Medco E&P Indonesia\Documents\VISCODE\scrape_p3dn_fix\cacert-2025-11-04.pem"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest"
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("auto_scraper")

def _verify_ssl():
    return CA_BUNDLE_PATH if os.path.exists(CA_BUNDLE_PATH) else True

# =============================
# 1. SCRAPING DATA KEMENPERIN (BMP)
# =============================
def scrape_kemenperin_bmp():
    url = "https://tkdn.kemenperin.go.id/export_excel_bmp.php"
    session = requests.Session()
    session.headers.update(HEADERS)
    
    try:
        session.get("https://tkdn.kemenperin.go.id/rekap.php", verify=_verify_ssl(), timeout=20)
        logger.info("Mengunduh data Kemenperin (BMP)...")
        resp = session.get(url, verify=_verify_ssl(), timeout=120)
        resp.raise_for_status()

        ctype = resp.headers.get("content-type", "").lower()
        if "html" in ctype or resp.content.lstrip().startswith(b"<"):
            dfs = pd.read_html(StringIO(resp.text), flavor="lxml")
            df = dfs[0] if dfs else pd.DataFrame()
        else:
            try:
                df = pd.read_excel(BytesIO(resp.content), engine="openpyxl")
            except Exception:
                df = pd.read_excel(BytesIO(resp.content), engine="xlrd")

        if not df.empty:
            if "No" not in df.columns:
                df.insert(0, "No", range(1, len(df) + 1))
            df.to_excel(OUTPUT_DIR / "bmp.xlsx", index=False, engine="openpyxl")
            df.to_csv(OUTPUT_DIR / "bmp.csv", index=False, encoding="utf-8-sig")
            logger.info(f"Sukses BMP: {len(df)} baris disimpan.")
    except Exception as e:
        logger.error(f"Gagal memproses data BMP: {e}")

# =============================
# 2. SCRAPING DATA ESDM (APDN BARANG & JASA)
# =============================
def scrape_apdn_esdm(name, url):
    try:
        logger.info(f"Mengambil data ESDM ({name}) dari API...")
        payload = {"start": 0, "length": 50000}
        
        resp = requests.post(url, headers=HEADERS, data=payload, timeout=120)
        resp.raise_for_status()
        
        data_json = resp.json()
        rows = data_json.get("data", [])
        
        if not rows:
            logger.warning(f"Data ESDM ({name}) kosong.")
            return

        df = pd.DataFrame(rows)
        if "No" not in df.columns:
            df.insert(0, "No", range(1, len(df) + 1))

        df.to_excel(OUTPUT_DIR / f"{name}.xlsx", index=False, engine="openpyxl")
        df.to_csv(OUTPUT_DIR / f"{name}.csv", index=False, encoding="utf-8-sig")
        logger.info(f"Sukses ESDM ({name}): {len(df)} baris disimpan.")

    except Exception as e:
        logger.error(f"Gagal mengambil data ESDM ({name}): {e}")

# =============================
# 3. GENERATE OTOMATIS KE INDEX.HTML
# =============================
def generate_web_page():
    logger.info("Membuat halaman web (index.html) dari file Excel...")
    
    # Gabungkan atau baca file utama yang ingin ditampilkan ke web (contoh: apdn_barang atau gabungan)
    # Di sini kita ambil contoh apdn_barang.xlsx (atau ubah sesuai kebutuhan file utama Anda)
    target_excel = OUTPUT_DIR / "apdn_barang.xlsx"
    
    if target_excel.exists():
        df = pd.read_excel(target_excel)
        # Ambil sampel atau keseluruhan (batasi jika terlalu besar agar ringan dibaca AI, misal 2000 baris pertama)
        html_table = df.head(3000).to_html(index=False, classes="table table-striped table-bordered")
    else:
        html_table = "<p>Data belum tersedia. Menunggu proses scraping berikutnya.</p>"

    html_content = f"""
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <title>Database TKDN & APDN Terupdate</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
</head>
<body class="container py-4">
    <h1 class="mb-4">Data Rekapitulasi TKDN & APDN Otomatis</h1>
    <p class="text-muted">Pembaruan terakhir: Otomatis 1x24 Jam via GitHub Actions</p>
    <div class="table-responsive">
        {html_table}
    </div>
</body>
</html>
"""

    # Simpan sebagai index.html tepat di root folder utama
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html_content)
    logger.info("File index.html berhasil diperbarui!")

# =============================
# EKSEKUSI UTAMA
# =============================
if __name__ == "__main__":
    logger.info("=== MEMULAI PIPELINE SCRAPING HARIAN ===")
    
    # 1. Jalankan Scraping Kemenperin
    scrape_kemenperin_bmp()
    time.sleep(2)
    
    # 2. Jalankan Scraping ESDM Jasa & Barang
    scrape_apdn_esdm("apdn_jasa", "https://www.esdm.go.id/apdn/Jasa/get")
    time.sleep(2)
    scrape_apdn_esdm("apdn_barang", "https://www.esdm.go.id/apdn/Barang/get")
    time.sleep(2)
    
    # 3. Generate file HTML untuk web publik GitHub Pages
    generate_web_page()
    logger.info("=== PIPELINE SELESAI ===")
