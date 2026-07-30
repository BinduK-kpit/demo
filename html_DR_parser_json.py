import json

# -----------------------------
# LOAD INPUT JSON
# -----------------------------
with open("html_DR_tags.json", "r") as f:
    data = json.load(f)

headers = []
rows = []

# -----------------------------
# PARSE THEAD + TR
# -----------------------------
for item in data:
    tag = item.get("tag")
    lines = item.get("lines", [])

    # clean lines
    lines = [line.strip() for line in lines if line.strip()]

    #  extract header
    if tag == "thead" and not headers:
        headers = lines

    #  extract rows
    elif tag == "tr":
        if lines:
            rows.append(lines)

# -----------------------------
# SAVE NEW JSON
# -----------------------------
output = {
    "headers": headers,
    "rows": rows
}

with open("structured_DR_table.json", "w") as f:
    json.dump(output, f, indent=4)

print(f" Created structured JSON with {len(rows)} rows")

