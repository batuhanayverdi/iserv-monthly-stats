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
            number = int(span.text.replace(".", "").strip())
            results[key] = number

# users_per_school hesapla (float, yuvarlama yok)
results["users_per_school"] = results["users"] / results["schools"]

# Excel dosyasını oku veya oluştur
try:
    df = pd.read_excel(EXCEL_PATH)
except FileNotFoundError:
    df = pd.DataFrame(columns=["month", "schools", "authorities", "users", "users_per_school"])

# Aynı ay varsa sil (datetime karşılaştırması için normalize et)
df["month"] = pd.to_datetime(df["month"]).dt.normalize()
current_month = results["month"].replace(tzinfo=None)
df = df[df["month"] != current_month]

# Yeni veriyi ekle
df = pd.concat([df, pd.DataFrame([results])], ignore_index=True)

# 🔧 Sütun sırasını garantiye al (özellikle users_per_school = F olması için)
df = df[["month", "schools", "authorities", "users", "users_per_school"]]

# Excel’e yaz
df.to_excel(EXCEL_PATH, index=False)

print("✅ Data updated successfully.")
