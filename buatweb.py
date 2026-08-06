import pandas as pd

df = pd.read_excel("result_apdn_kb/apdn_jasa.xlsx")

html_table = df.to_html(index=False, classes="table table-striped table-bordered")

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
    <p class="text-muted">Pembaruan terakhir: Setiap 1x24 Jam</p>
    <div class="table-responsive">
        {html_table}
    </div>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
  f.write(html_content)

print("File index.html berhasil diperbarui dengan tabel data!")