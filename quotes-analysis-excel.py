import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font
import pandas as pd

quotes = []
authors = []

for page in range(1, 4):
    url = f"http://quotes.toscrape.com/page/{page}/"

    response = requests.get(url)
    response.encoding = "utf-8"

    soup = BeautifulSoup(response.text, "html.parser")

    quote = soup.find_all("span", class_="text")
    author = soup.find_all("small", class_="author")

    for i in range(len(quote)):
        quotes.append(quote[i].text)
        authors.append(author[i].text)

print(f"Total quotes: {len(quotes)}")

dp = pd.DataFrame({"Quote": quotes, "Author": authors})
dp.drop_duplicates(inplace=True)

author_counts = dp["Author"].value_counts()
print(author_counts)

with pd.ExcelWriter("books_quotes.xlsx", engine="openpyxl") as writer:
    dp.to_excel(writer, index=False, sheet_name="Books")
    worksheet = writer.sheets["Books"]
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
print("excel file ready")