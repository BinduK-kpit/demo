import requests
import json
from lxml import html
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv

load_dotenv()
# -----------------------------
# CONFIG
# -----------------------------
url = os.getenv("CONFLUENCE_DR_URL")
api_token=os.getenv("CONFLUENCE_API_TOKEN")

headers = {
    'User-Agent': 'Mozilla/5.0',
    'Authorization': f'Bearer {api_token}'
}

# -----------------------------
# FETCH DATA
# -----------------------------
response = requests.get(url, headers=headers)
data = response.json()

print(type(data))

# -----------------------------
#  EXTRACT RAW HTML FROM PAGE
# -----------------------------
html_content = data.get("body", {}).get("view", {}).get("value", "")

# -----------------------------
#  BEAUTIFULSOUP - FULL TEXT
# -----------------------------
if html_content:
    soup = BeautifulSoup(html_content, "html.parser")

    text = soup.get_text(separator="\n")

    # Save raw text
    with open("page_text.txt", "w") as f:
        f.write(text)

    print(" Extracted full page text")

# -----------------------------
#  LXML PARSING (OPTIONAL STRUCTURE)
# -----------------------------
parsed = []

if html_content:
    tree = html.fromstring(html_content)

    # Example: extract ALL tags
    for element in tree.iter():
        record = {
            "tag": element.tag,
            "text": element.text_content().strip() if element.text_content() else None,
            "attributes": element.attrib
        }
        parsed.append(record)

processed = []

for element in parsed:

    raw_text = element.get("text")

    if raw_text:
        #  split by newline and clean
        lines = [line.strip() for line in raw_text.split("\n") if line.strip()]
    else:
        lines = []

    processed.append({
        "tag": element.get("tag"),
        "lines": lines
    })

# -----------------------------
# SAVE TAG-LEVEL JSON
# -----------------------------
with open("html_DR_tags.json", "w") as f:
    json.dump(processed, f, indent=4)

print(f" Saved {len(processed)} HTML elements")

# -----------------------------
#  OPTIONAL: EXTRACT TABLE DATA (if any)
# -----------------------------
tables = tree.xpath("//table")

if tables:
    print(f" Found {len(tables)} tables")

print(" Done")

