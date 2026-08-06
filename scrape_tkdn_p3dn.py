import os
import time
import logging
import requests
import pandas as pd
from pathlib import Path
from io import BytesIO, StringIO

# =============================
# KONSTANTA & CONFIG
# =============================
OUTPUT_DIR = Path("result_copilot_kb")
CA_BUNDLE_PATH = r"C:\Users\90006790\OneDrive - PT. Medco E&P Indonesia\Documents\VISCODE\scrape_p3dn_fix\cacert-2025-11-04.pem"

TARGET_FILES = {
    "tkdn_lvi_2026": "https://tkdn.kemenperin.go.id/export_excel.php?thn=2LtIAq-S_CL6ShnR120ukm23GZ4e8eNd5rq2cOwvPnM", 
    "tkdn_lvi_2025": "https://tkdn.kemenperin.go.id/export_excel.php?thn=2sOoXJh1Pq6y1tkg5oA7MV6Ia96M2Hfla6_ABTCB2u8",
    "tkdn_lvi_2024": "https://tkdn.kemenperin.go.id/export_excel.php?thn=_XO3tSoRY4lHDtw2EKfSt_-1gNj9Q-dwKGgBz5nb8EM",
    "bmp": "https://tkdn.kemenperin.go.id/export_excel_bmp.php."
}

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0 Safari/537.36"}

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("copilot_data_prep")

def _verify_ssl():
    return CA_BUNDLE_PATH if os.path.exists(CA_BUNDLE_PATH) else True

def download_and_save(name, url):
    if "..." in url:
        logger.warning(f"URL untuk {name} belum diisi dengan benar. Lewati...")
        return

    try:
        logger.info(f"Mengunduh data '{name}'...")
        resp = requests.get(url, headers=HEADERS, verify=_verify_ssl(), timeout=60)
        resp.raise_for_status()

        ctype = resp.headers.get("content-type", "").lower()

        # Smart Parsing: Menangani format HTML wrapper atau Excel biner asli
        if "html" in ctype or resp.content.lstrip().startswith(b"<"):
            dfs = pd.read_html(StringIO(resp.text), flavor="lxml")
            df = dfs[0] if dfs else pd.DataFrame()
        else:
            try:
                df = pd.read_excel(BytesIO(resp.content), engine="openpyxl")
            except Exception:
                df = pd.read_excel(BytesIO(resp.content), engine="xlrd")

        if df.empty:
            logger.warning(f"Data untuk {name} kosong.")
            return

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        excel_path = OUTPUT_DIR / f"{name}.xlsx"
        csv_path = OUTPUT_DIR / f"{name}.csv"

        if "No" not in df.columns:
            df.insert(0, "No", range(1, len(df) + 1))

        # Simpan ke Excel & CSV
        df.to_excel(excel_path, index=False, engine="openpyxl")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        
        logger.info(f"Sukses! Disimpan ke: {excel_path.name} & {csv_path.name} ({len(df)} baris)")

    except Exception:
        logger.exception(f"Gagal memproses data untuk {name}")

def main():
    for name, url in TARGET_FILES.items():
        download_and_save(name, url)
        time.sleep(1)

if __name__ == "__main__":
    main()