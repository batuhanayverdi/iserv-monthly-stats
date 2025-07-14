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

results = {"month": datetime.today().strftime("%Y-%m")}

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

# users_per_school hesapla
results["users_per_school"] = round(results["users"] / results["schools"])

# Excel dosyasını güncelle
try:
    df = pd.read_excel(EXCEL_PATH)
except FileNotFoundError:
    df = pd.DataFrame(columns=["month", "schools", "authorities", "users", "users_per_school"])

# Aynı ay varsa sil
df = df[df["month"] != results["month"]]

# Yeni veriyi ekle
df = pd.concat([df, pd.DataFrame([results])], ignore_index=True)
df.to_excel(EXCEL_PATH, index=False)

print("✅ Data updated successfully.")
