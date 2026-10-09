"""
GGS (Gloucester Garden Supplies): flatten the nested JSON export into a CSV.

Why: Excel for the web has no Power Query, so it cannot open nested JSON.
pandas.json_normalize turns each record into one flat row.

I ran this in Google Colab (uploading the file with google.colab.files).
It also runs locally:  python ggs_flatten_json.py
"""
import json
import pandas as pd

SOURCE = "../data/01_client_submissions/GGS_GloucesterGarden_ERP_export_2025.json"
OUTPUT = "GGS_records.csv"

with open(SOURCE, encoding="utf-8") as f:
    data = json.load(f)

df = pd.json_normalize(data["records"])

# Intake check: how many rows per region? (this is how the Ireland rows showed up)
print(df["ship_to_region"].value_counts())

df.to_csv(OUTPUT, index=False)
print(f"Saved {len(df)} rows to {OUTPUT}")
