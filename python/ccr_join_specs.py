"""
CCR (Clifton Coffee Roasters): join sales units to per-unit packaging specs.

Sales file:  one row per SKU, nation and channel, with H1 and H2 units.
Specs file:  one row per SKU and packaging component, with grams per unit.

One sales row matches several spec rows (many-to-many by SKU), so an Excel
lookup would return only the first component and silently get it wrong.

Key decisions:
  - LEFT join, not inner: SKUs sold with no spec (CCR-0038, 0044, 0045) must
    stay visible as blanks to query, not disappear.
  - Check for duplicate spec versions: CCR-0002 had a 2024 and a 2025 spec,
    which would count every component twice.
  - Unpivot (melt) H1/H2 units into one column with a period label.
  - weight (g) = units x grams per unit.

Run:  python ccr_join_specs.py
"""
import pandas as pd

SALES = "../data/01_client_submissions/CCR_CliftonCoffee_sales_2025.xlsx"
SPECS = "../data/01_client_submissions/CCR_CliftonCoffee_packaging_specs.csv"
OUTPUT = "CCR_joined_long.xlsx"

sales = pd.read_excel(SALES)
specs = pd.read_csv(SPECS)

m = sales.merge(specs, on="SKU", how="left")

long = m.melt(
    id_vars=["SKU", "Product", "Channel", "Nation", "Component",
             "Material", "Class", "Unit weight (g)", "Spec date"],
    value_vars=["H1 Units", "H2 Units"],
    var_name="raw_period",
    value_name="units",
)

long["raw_weight"] = long["units"] * long["Unit weight (g)"]

# Check 1: SKUs that were sold but have no spec (blank component after the left join)
print("SKUs with no spec:", long[long["Component"].isna()]["SKU"].unique())

# Check 2: SKUs with more than one spec version (would double-count weight)
print("SKUs with more than one spec date:")
print(long.groupby("SKU")["Spec date"].nunique().loc[lambda s: s > 1])

long.to_excel(OUTPUT, index=False)
print(f"Saved {len(long)} rows to {OUTPUT}")
