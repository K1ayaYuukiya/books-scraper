import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from pathlib import Path
from openpyxl.styles import Font, PatternFill

RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

books = []

for page in range(1, 51):
    url = f"https://books.toscrape.com/catalogue/page-{page}.html"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")

    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        price_text = card.select_one("p.price_color").text
        rating_word = card.select_one("p.star-rating")["class"][1]

        books.append({
            "Title": link["title"],
            "Price (£)": float(price_text.replace("£", "").replace("Â", "")),
            "Rating": RATINGS.get(rating_word, 0),
            "Availability": card.select_one("p.availability").text.strip(),
            "URL": urljoin(url, link["href"]),
        })

    print(f"Страница {page}/50 готова, всего книг: {len(books)}")
    time.sleep(0.5)

df = pd.DataFrame(books)

output = Path(__file__).parent / "books.xlsx"
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Books")
    sheet = writer.sheets["Books"]
    widths = {"A": 60, "B": 12, "C": 10, "D": 15, "E": 80}
    for col, width in widths.items():
        sheet.column_dimensions[col].width = width
    sheet.freeze_panes = "A2"
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet["B"][1:]:
        cell.number_format = "0.00"
    sheet.auto_filter.ref = sheet.dimensions

print("Готово! Файл books.xlsx сохранён")