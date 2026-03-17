import requests
import time
from openpyxl import load_workbook

def query_wikidata(sparql_query, limit=1000, offset=0):
    url = "https://query.wikidata.org/sparql"
    headers = {
        "User-Agent": "MyBot/1.0 (your-email@example.com)",
        "Accept": "application/json"
    }

    if "LIMIT" not in sparql_query.upper():
        sparql_query += f" LIMIT {limit} OFFSET {offset}"

    params = {
        "query": sparql_query,
        "format": "json"
    }
    response = requests.get(url, params=params, headers=headers, timeout=30)

    if response.status_code == 200:
        data = response.json()
        return data["results"]["bindings"]
    else:
        print(f"Ошибка {response.status_code}: {response.text[:200]}")
        return []


class_qid = 'Q198'
query = f"""
SELECT DISTINCT ?item ?itemLabel ?start ?end WHERE {{
  ?item wdt:P31/wdt:P279* wd:{class_qid} .
  OPTIONAL {{ ?item wdt:P580 ?start . }}
  OPTIONAL {{ ?item wdt:P582 ?end . }}
  OPTIONAL {{ ?item wdt:P585 ?point . }}
  SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
}}
"""

start = time.perf_counter()
results = query_wikidata(query, limit=100)


fn = "example.xlsx"

wb = load_workbook(fn)
ws = wb['Лист1']
s = ['id', 'Название', 'Начало', 'Конец', 'url']
for i in range(1, 6):
    ws.cell(row=1, column=i).value = s[i-1]


for i, res in enumerate(results, 2):
    item_url = res["item"]["value"]
    item_qid = item_url.split("/")[-1]
    label = res.get("itemLabel", {}).get("value", "нет метки")

    start_date = res.get("start", {}).get("value", "нет данных")
    end_date = res.get("end", {}).get("value", "нет данных")
    point = res.get("point", {}).get("value", "нет данных")

    ws.cell(row=i, column=1).value = item_qid
    ws.cell(row=i, column=2).value = label
    if start_date == "нет данных":
        ws.cell(row=i, column=3).value = point
    else:
        ws.cell(row=i, column=3).value = start_date
    ws.cell(row=i, column=4).value = end_date
    ws.cell(row=i, column=5).value = item_url

wb.save(fn)
wb.close()

end = time.perf_counter()
print(f"Количество найденных элементов: {len(results)}")
print(f"Время выполнения: {end - start:.6f} секунд")