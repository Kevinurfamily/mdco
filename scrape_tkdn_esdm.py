import requests
import pandas as pd
from pathlib import Path

OUTPUT_DIR = Path("result_apdn_kb")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Endpoint POST yang Anda temukan di Network
URL_JASA_API = "https://www.esdm.go.id/apdn/Jasa/get"
URL_BARANG_API = "https://www.esdm.go.id/apdn/Barang/get"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest" # Penanda penting bahwa ini adalah request AJAX
}

def fetch_apdn_data(name, url):
    try:
        print(f"Mengambil data {name} dari API...")
        
        # Data/Payload yang dikirim ke server (biasanya DataTables mengirim parameter length/start)
        # Kita set length besar (misal 50000) agar server mengembalikan semua data sekaligus
        payload = {
            "start": 0,
            "length": 50000 
        }

        resp = requests.post(url, headers=HEADERS, data=payload, timeout=120)
        resp.raise_for_status()
        
        # Ubah respons JSON dari server
        data_json = resp.json()
        
        # DataTables biasanya menyimpan array data di dalam key 'data'
        rows = data_json.get("data", [])
        
        if not rows:
            print(f"Data {name} kosong atau format JSON berbeda.")
            return

        df = pd.DataFrame(rows)
        
        # Simpan ke Excel & CSV
        if "No" not in df.columns:
            df.insert(0, "No", range(1, len(df) + 1))

        df.to_excel(OUTPUT_DIR / f"{name}.xlsx", index=False, engine="openpyxl")
        df.to_csv(OUTPUT_DIR / f"{name}.csv", index=False, encoding="utf-8-sig")
        
        print(f"Sukses! Berhasil menyimpan {len(df)} baris ke {name}.xlsx & {name}.csv")

    except requests.exceptions.Timeout:
        print(
        f"Gagal {name}: Server ESDM terlalu lama merespons (Timeout). Coba"
        " jalankan ulang beberapa saat lagi."
    )
    except Exception as e:
        print(f"Gagal mengambil data {name}: {e}")

if __name__ == "__main__":
    fetch_apdn_data("apdn_jasa", URL_JASA_API)
    fetch_apdn_data("apdn_barang", URL_BARANG_API)