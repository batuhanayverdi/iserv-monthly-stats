import pandas as pd
from datetime import datetime
import requests
from bs4 import BeautifulSoup

EXCEL_PATH = "iserv_stats.xlsx"

# Sayfayı çek
url = "https://www.iserv.eu/"
response = requests.get(url)
soup = BeautifulSoup(response.text, "html.parser")

# Sayıları çek (gerekirse class name değiştirilebilir)
numbers = soup.find_all("span", class_="counter")
schools      = int(numbers[0].text.replace(".", "").strip())
users        = int(numbers[2].text.replace(".", "").strip())
authorities  = int(numbers[3].text.replace(".", "").strip())
users_per_school = round(users / schools)

# Tarih = ilgili ayın ilk günü
today = datetime.today().replace(day=1)

# Excel’i oku ve yeni satır ekle
df = pd.read_excel(EXCEL_PATH)

# Eğer zaten bu ay ekliyse tekrar yazma
if today.strftime("%Y-%m-%d") not in df["month"].astype(str).values:
    new_row = {
        "month": today,
        "schools": schools,
        "authorities": authorities,
        "users": users,
        "users_per_school": users_per_school
    }
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_PATH, index=False)
