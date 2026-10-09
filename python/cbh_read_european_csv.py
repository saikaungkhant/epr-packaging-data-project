"""
CBH (Cardiff Bay Homewares): read a European-format CSV correctly.

The file uses:
  - semicolons as the column separator
  - Windows-1252 encoding (so Cafe/Creme accents read properly)
  - a comma as the decimal point      (1,5   means 1.5)
  - a dot as the thousands separator  (12.500 means 12500)

Getting decimal and thousands the wrong way round makes weights
10 to 1,000 times wrong, so they are set explicitly. The date column is
read as text, otherwise the dot-as-thousands rule turns 30.06.2025 into a number.

Run:  python cbh_read_european_csv.py
"""
import pandas as pd

SOURCE = "../data/01_client_submissions/CBH_CardiffBayHomewares_2025.csv"
OUTPUT = "CBH_CardiffBayHomewares_2025_separated_data.xlsx"

df = pd.read_csv(
    SOURCE,
    sep=";",
    encoding="cp1252",
    decimal=",",
    thousands=".",
    dtype={"Period_end": str},   # keep 30.06.2025 as text, or the "." rule mangles it
)

# Drop the client's TOTAL row (it has no SKU) so it is not counted as data
df = df[df["SKU"].notna()]

print(df.head())
print(df.dtypes)

df.to_excel(OUTPUT, index=False)
print(f"Saved {len(df)} rows to {OUTPUT}")
