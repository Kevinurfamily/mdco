import io
import logging
from pathlib import Path
import pandas as pd
import requests

OUTPUT_DIR = Path("result_copilot_kb")
CA_BUNDLE_PATH = r"C:\Users\90006790\OneDrive - PT. Medco E&P Indonesia\Documents\VISCODE\scrape_p3dn_fix\cacert-2025-11-04.pem"

TARGET_FILES = {
    "bmp": "https://tkdn.kemenperin.go.id/export_excel_bmp.php"
}

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0 Safari/537.36"}

def _verify_ssl():
    return CA_BUNDLE_PATH if Path(CA_BUNDLE_PATH).exists() else True

def download_bmp_with_session():
    session = requests.Session()
    session.headers.update(HEADERS)
    
    # 1. Kunjungi halaman rekap utama dulu untuk inisialisasi session/cookie
    try:
        session.get("https://tkdn.kemenperin.go.id/rekap.php", verify=_verify_ssl(), timeout=20)
    except Exception:
        pass

    # 2. Lakukan download file BMP
    url = TARGET_FILES["bmp"]
    try:
        print(f"Mengunduh data 'bmp' dari: {url}")
        # Timeout dinaikkan ke 120 detik karena file BMP biasanya besar
        resp = session.get(url, verify=_verify_ssl(), timeout=120)
        resp.raise_for_status()

        ctype = resp.headers.get("content-type", "").lower()

        if "html" in ctype or resp.content.lstrip().startswith(b"<"):
            dfs = pd.read_html(io.StringIO(resp.text), flavor="lxml")
            df = dfs[0] if dfs else pd.DataFrame()
        else:
            try:
                df = pd.read_excel(io.BytesIO(resp.content), engine="openpyxl")
            except Exception:
                df = pd.read_excel(io.BytesIO(resp.content), engine="xlrd")

        if df.empty:
            print("Data BMP kosong.")
            return

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        if "No" not in df.columns:
            df.insert(0, "No", range(1, len(df) + 1))

        df.to_excel(OUTPUT_DIR / "bmp.xlsx", index=False, engine="openpyxl")
        df.to_csv(OUTPUT_DIR / "bmp.csv", index=False, encoding="utf-8-sig")
        print(f"Sukses! Disimpan ke bmp.xlsx & bmp.csv ({len(df)} baris)")

    except Exception as e:
        print(f"Gagal memproses data BMP: {e}")

if __name__ == "__main__":
    download_bmp_with_session()