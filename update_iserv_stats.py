import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime

EXCEL_PATH = "iserv_stats.xlsx"
URL = "https://iserv.de/"

response = requests.get(URL)
soup = BeautifulSoup(response.content, "html.parser")

labels_map = {
    "überzeugte Schulen": "schools",
    "Benutzer(innen)": "users",
    "überzeugte Schulträger": "authorities"
}

# Aylık veri: Ayın 1’i ve saat 00:00:00 formatında
results = {"month": datetime.today().replace(day=1, hour=0, minute=0, second=0, microsecond=0)}

# Verileri çek
for span in soup.find_all("span", class_="text-iserv-blue-024 font-bold"):
    parent = span.find_parent()
    label_span = parent.find("span", class_="text-iserv-neutral-035")
    if label_span:
        label = label_span.text.strip()
        key = labels_map.get(label)
        if key:
            number = int(span.text.replace(".", "").strip())  # Sayıları doğru al
            results[key] = number

# users_per_school hesapla (float, yuvarlama yok)
results["users_per_school"] = results["users"] / results["schools"]

# Excel dosyasını oku veya oluştur
try:
    df = pd.read_excel(EXCEL_PATH)
except FileNotFoundError:
    df = pd.DataFrame(columns=[
        "month", "schools", "authorities", "users", "users_per_school",
        "yoy_schools", "yoy_authorities", "yoy_users"
    ])

# Tarihleri normalize et, aynı ay varsa sil
df["month"] = pd.to_datetime(df["month"]).dt.normalize()
current_month = results["month"].replace(tzinfo=None)
df = df[df["month"] != current_month]

# Yeni veriyi ekle
df = pd.concat([df, pd.DataFrame([results])], ignore_index=True)

# YoY hesapla (sadece 2025 ve sonrası için)
df["yoy_schools"] = pd.NA
df["yoy_authorities"] = pd.NA
df["yoy_users"] = pd.NA

for i, row in df.iterrows():
    current_month = pd.to_datetime(row["month"])
    if current_month.year >= 2025:
        prev_year_row = df[df["month"] == current_month.replace(year=current_month.year - 1)]
        if not prev_year_row.empty:
            try:
                df.at[i, "yoy_schools"] = ((row["schools"] / prev_year_row["schools"].values[0]) - 1)
                df.at[i, "yoy_authorities"] = ((row["authorities"] / prev_year_row["authorities"].values[0]) - 1)
                df.at[i, "yoy_users"] = ((row["users"] / prev_year_row["users"].values[0]) - 1)
            except Exception as e:
                print(f"YoY hesaplama hatası ({row['month']}): {e}")

# Sütun sırasını sabitle
df = df[[
    "month", "schools", "authorities", "users", "users_per_school",
    "yoy_schools", "yoy_authorities", "yoy_users"
]]

# Excel'e yaz
with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
    df = df.astype({
        "schools": int,
        "authorities": int,
        "users": int,
        "users_per_school": float
    })
    df.to_excel(writer, index=False)

print("✅ Data updated successfully.")
