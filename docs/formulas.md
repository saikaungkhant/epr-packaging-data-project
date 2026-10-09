# Every formula, and why

This part lists every formula in the order I built them. For each one: the formula, what it does, and why I used it that way. Column letters are the ones in my Raw_ALL sheet:

| Col | Field | Col | Field | Col | Field |
|---|---|---|---|---|---|
| A | client_id | L | note | W | dup_key |
| B | sku | M | clean_material | X | flag_duplicate |
| C | component | N | clean_class | Y | g_per_unit |
| D | raw_material | O | clean_type | Z | grp_key |
| E | raw_class | P | clean_nation | AA | ratio |
| F | raw_type | Q | clean_period | AB | flag_unit_weight |
| G | raw_nation | R | assumption | AC | override_kg |
| H | raw_period | S | weight_kg | AD | exclude_reason |
| I | units | T | flag_weight | AE | final_kg |
| J | raw_weight | U | flag_code | AF | sub_key |
| K | raw_weight_unit | V | flag_hdc | | |

## A. Checking counts (Module 2)

### Rows per client

```
=COUNTIF(Raw_All!A:A,"AVD")
```

**What it does:** counts the rows for one client. I put one per client next to the expected count, with a difference column.

**Why:** the total row count was wrong (10,266, then 20,135). A per-client table showed instantly that only BBO was off (+5). **Always reconcile per group**, because a total can hide errors.

## B. Mapping and cleaning (Module 3)

### List the distinct raw values

```
=UNIQUE(Raw_All!D2:D10249)
```

**What it does:** lists every different spelling in a column once.

**Why:** you cannot build a mapping table until you know every spelling the clients used. A "0" in the result means a blank cell.

### Mappings key (on the Mappings sheet, column F)

```
=A2&"|"&B2
```

**What it does:** joins field and raw_value into one key, for example `material|Plastik` or `class|HH - Secondary`.

**Why:** the same text can mean different things in different columns ("HH Drinks - Primary" is HDC as a type but P1 as a class). Putting the field name in the key keeps them apart, and gives a single column to look up.

### clean_material (M)

```
=IF(D2="","",XLOOKUP("material|"&TRIM(D2),Mappings!$F$2:$F$122,Mappings!$C$2:$C$122,"UNMAPPED"))
```

**What it does:** builds the same key from the raw value and looks up the official code.

**Why each part:**

- `IF(D2="","",...)` keeps blanks blank, instead of turning them into "UNMAPPED".
- `TRIM` removes trailing spaces, so "Plastic  " finds the same entry as "Plastic".
- `"UNMAPPED"` as the not-found value makes any forgotten spelling **loud** instead of silently blank.

**Tip I learned the hard way:** use whole columns, `Mappings!F:F` and `Mappings!C:C`. With a fixed range like $F$2:$F$122, a new mapping row at row 123 is ignored (the "2025-H2" problem). The formulas below are shown with whole columns, which is how they should all be.

### clean_class (N)

```
=IF(E2<>"",XLOOKUP("class|"&TRIM(E2),Mappings!F:F,Mappings!C:C,"UNMAPPED"),IF(A2="MDC",XLOOKUP("class|"&TRIM(F2),Mappings!F:F,Mappings!C:C,"UNMAPPED"),""))
```

**What it does:** normally looks up raw_class. For MDC, which has no class column, it looks up the type label instead ("NHH - Shipment" gives P3).

**Why:** this is Rule 2. MDC hid the class inside its type label.

**The bug this formula had:** on the MDC rows, a copy of this formula pointed at E5, F5 and A5 (an AVD row) instead of its own row, so every MDC row became P1. Always check a formula refers to **its own row**.

### clean_type (O)

```
=IF(AND(A2="CCR",N2="P3"),"NH",IF(F2<>"",XLOOKUP("type|"&TRIM(F2),Mappings!F:F,Mappings!C:C,"UNMAPPED"),IF(OR(A2="BBO",A2="CBH"),IF(OR(N2="P1",N2="P2"),"HH",IF(OR(N2="P3",N2="P4"),"NH","")),"")))
```

**What it does, in order:**

1. A CCR shipping case (P3) is always NH.
2. Otherwise, if there is a raw type, look it up.
3. Otherwise, for BBO and CBH (no type given), assume from the class: P1/P2 is HH, P3/P4 is NH.

**Why the order matters:** the CCR rule has to come first. If the lookup ran first, a shipping case sold through "Retail" would become HH.

### clean_nation (P)

```
=IF(G2="","",XLOOKUP("nation|"&TRIM(G2),Mappings!F:F,Mappings!C:C,"UNMAPPED"))
```

**What it does:** maps Eng, England, ENG, N. Ireland, Nor, IRL and so on to EN, SC, WS, NI or OUT_OF_SCOPE.

**Why:** IRL maps to OUT_OF_SCOPE rather than blank, so the Ireland rows stay visible and are excluded **on purpose** later. For MDC (no nation given) I entered EN, based on the client register (Rule 1).

### clean_period (Q)

```
=IF(H2="","",XLOOKUP("period|"&TRIM(H2),Mappings!F:F,Mappings!C:C,"UNMAPPED"))
```

**What it does:** maps H1 2025, 2025 H1, Jan-Jun 25, 45838 and so on to 2025-H1 or 2025-H2.

**Why whole columns:** this is the one that broke with 752 UNMAPPED, because "2025-H2" was missing from Mappings and the old fixed range would not have seen a new row anyway.

### assumption (R)

```
=IF(AND(F2="",OR(A2="BBO",A2="CBH")),"Type assumed: confirm with client","")
```

**What it does:** labels every row where the packaging type was assumed rather than given.

**Why:** an assumption should never be hidden. This column feeds the query log, so the client can confirm it.

### weight_kg (S)

```
=IFERROR(VALUE(IF(TRIM(K2)="g",SUBSTITUTE(TRIM(J2),",","."),SUBSTITUTE(SUBSTITUTE(TRIM(J2),"t",""),",","")))*SWITCH(TRIM(K2),"kg",1,"tonnes",1000,"t",1000,"g",0.001,"lbs",1/2.20462),"")
```

**What it does, inside out:**

1. `TRIM(J2)` removes spaces around the number.
2. If the unit is grams (CBH), the comma is a **decimal point**, so it becomes a dot.
3. Otherwise, remove the trailing "t" (SSN) and remove commas, which are **thousands separators** (BBO).
4. `VALUE` turns the cleaned text into a number.
5. `SWITCH` picks the multiplier for the unit: kg x1, tonnes x1000, g x0.001, lbs divided by 2.20462.
6. `IFERROR(...,"")` leaves "n/a" and "TBC" blank, because they are **queries, not zeros**.

**Why it looks like this:** the first version was simpler, but I tried to clean commas with Find & Replace and it wiped every comma from every formula. Moving the comma logic into the formula, depending on the unit, means the raw data never gets touched.

## C. Validation flags (Module 4)

### flag_weight (T)

```
=IF(S2="","MISSING WEIGHT",IF(S2<=0,"NEGATIVE OR ZERO",""))
```

**What it does:** flags weights that are missing, negative or zero.

**Why:** found 43 missing (AVD 12, SSN 8, BBO 5, CCR 18) and 5 negative SSN stock adjustments.

### flag_code (U)

```
=IF(OR(M2="UNMAPPED",N2="UNMAPPED",O2="UNMAPPED",P2="UNMAPPED",Q2="UNMAPPED"),"UNMAPPED",IF(P2="OUT_OF_SCOPE","OUT OF SCOPE",""))
```

**What it does:** one column that tells me if any of the five clean columns failed, or if the row is out of scope.

**Why:** one place to count from. `=COUNTIF(U:U,"UNMAPPED")` should always be 0.

### flag_hdc (V)

```
=IF(O2<>"HDC","",IF(ISNA(MATCH(M2,{"AL","GL","PL","ST"},0)),"INVALID HDC MATERIAL",IF(I2="","HDC NO UNITS","")))
```

**What it does:** only for HDC rows, checks the material is one of AL, GL, PL or ST, and that units are present.

**Why:** these are the two HDC rules. `MATCH(M2,{"AL","GL","PL","ST"},0)` looks for the material in a small list typed inside the formula, and ISNA is TRUE when it is not found. This caught MDC's milk cartons (paper as HDC).

### dup_key (W) and flag_duplicate (X)

```
=A2&"|"&B2&"|"&C2&"|"&D2&"|"&F2&"|"&G2&"|"&H2&"|"&I2&"|"&J2&"|"&L2
=IF(COUNTIF(W$2:W$10249,W2)>1,"DUPLICATE","")
```

**What it does:** joins every field that makes a row unique, then flags any key that appears more than once.

**Why units (I2) is in there:** without it, two different lines with the same weight looked like duplicates. With it: 12 flagged rows, the 6 real AVD pairs.

### g_per_unit (Y) and grp_key (Z)

```
=IF(AND(ISNUMBER(S2),ISNUMBER(I2)),IF(AND(S2>0,I2>0),S2*1000/I2,""),"")
=B2&"|"&C2
```

**What they do:** grams per unit for each row, and a key for the SKU and component group.

**Why the guards:** only calculate when both weight and units are real positive numbers, so no divide-by-zero errors.

### UnitCheck sheet: the median per group

```
A2: =UNIQUE(FILTER(Raw_All!Z2:Z10249,Raw_All!Y2:Y10249<>""))
B2: =IF(A2="","",MEDIAN(FILTER(Raw_All!$Y$2:$Y$10249,(Raw_All!$Z$2:$Z$10249=A2)*(Raw_All!$Y$2:$Y$10249<>""))))
```

**What it does:** lists each SKU and component group once, then calculates the median grams per unit for each group.

**Why median:** one x10 error drags an average so far that the correct rows look wrong. The median ignores one extreme value. **Why a separate sheet:** about 1,770 groups calculated once is much faster than a MEDIAN(FILTER) on every one of 11,000 rows.

**How the FILTER works:** multiplying two TRUE/FALSE conditions acts as AND. Only rows in this group AND with a value are kept.

### ratio (AA) and flag_unit_weight (AB)

```
=IF(Y2="","",Y2/XLOOKUP(Z2,UnitCheck!$A$2:$A$2000,UnitCheck!$B$2:$B$2000))
=IF(AA2="","",IF(OR(AA2>3,AA2<1/3),"CHECK UNIT WEIGHT",""))
```

**What they do:** compare each row's grams per unit with its group's median, and flag anything more than 3 times bigger or smaller.

**Why:** catches errors that pass every other check. A ratio near 0.001 means tonnes typed as kg. Near 10 means an extra zero. Found exactly 5 rows.

### Period completeness (Checks sheet)

```
=COUNTIFS(Raw_All!$A:$A,$A2,Raw_All!$Q:$Q,B$1)
```

**What it does:** with clients down column A and periods across row 1, counts rows for each client and period.

**Why:** a 0 shows a missing period. It confirmed SPF had no H2.

## D. Corrections (Module 5)

### exclude_reason (AD)

```
=IF(AND(A2="CCR",LEFT(L2,6)<>"rejoin"),"Replaced by CCR re-join (CCR-01/02)",IF(P2="OUT_OF_SCOPE","Republic of Ireland (GGS-01)",IF(T2="NEGATIVE OR ZERO","Stock adjustment (SSN-02)",IF(X2="DUPLICATE",IF(COUNTIF(W$2:W2,W2)>1,"Duplicate (AVD-03)",""),""))))
```

**What it does:** gives each row that should not be submitted a written reason, with the query ID.

**Why each part:**

- Old CCR rows (anything without the "rejoin" note) are replaced by the rebuilt CCR block.
- Ireland and stock adjustments are excluded as agreed with the clients.
- For duplicates, `COUNTIF(W$2:W2,W2)` uses an **expanding range** (the top is fixed, the bottom moves down). It counts how many times the key has appeared **so far**, so the first copy gets 1 (kept) and the second gets 2 (excluded).

**Why a formula instead of deleting:** nothing is lost, every exclusion has a reason, and it can be reviewed or undone.

### final_kg (AE)

```
=IF(AD2<>"","",IF(AC2<>"",AC2,S2))
```

**What it does:** blank if excluded; otherwise the override (AC) if there is one; otherwise the converted weight (S).

**Why:** the raw weight and the flag stay untouched, and the correction lives in its own column. This is the only column the submission adds up.

### Checks after the SKU accident

```
=COUNTIF(B:B,"AVD-0002")
=COUNTIFS(A:A,"<>AVD",B:B,"AVD*")
=SUMPRODUCT((B2:B10249<>"")*(LEFT(B2:B10249,3)<>A2:A10249))
```

**What they do:** the first found 72 instead of 46. The second counts AVD SKUs sitting on non-AVD rows (`*` is a wildcard). The third counts any row whose SKU prefix does not match its client.

**Why:** after a repair, you prove it with a count that should be 0.

### Finding missing rows (anti-join)

```
=COUNTIFS(Raw_ALL!B:B,B2,Raw_ALL!C:C,C2,Raw_ALL!P:P,M2,Raw_ALL!Q:Q,N2)
```

**What it does:** on a list of the rows that **should** exist, counts how many times each one appears in Raw_ALL by SKU, component, nation and period. A 0 means the row is missing.

**Why:** this found the exact 4 BBO rows lost to the off-by-one deletion, and later the 3 CBH rows lost to the paste.

### Diagnosing the 752 UNMAPPED

```
=LEN(Raw_ALL!H11895)
=COUNTIF(Mappings!F:F,"period|2025-H2")
```

**What they do:** LEN showed 7 characters, so the text was clean. COUNTIF showed 0, so the mapping row did not exist.

**Why:** test one layer at a time. First the data, then the dictionary.

## E. Building the submission (Module 6)

### sub_key (Raw_ALL, AF)

```
=IF(AD2<>"","",A2&"|"&Q2&"|"&XLOOKUP(A2,Clients!A:A,Clients!C:C)&"|"&O2&"|"&N2&"|"&M2&"|"&P2)
```

**What it does:** glues client, period, activity, type, class, material and nation into one label, like `AVD|2025-H1|SO|HDC|P1|AL|EN`. Excluded rows get blank.

**Why:** every row that belongs to the same submission line gets the same key. Activity comes from the small Clients sheet.

### The Submission sheet: 5 spilling formulas

```
L2: =SORT(UNIQUE(FILTER(Raw_ALL!AF2:AF20000,Raw_ALL!AF2:AF20000<>"")))
B2: =TEXTSPLIT(TEXTJOIN(";",,L2#),"|",";")
A2: =XLOOKUP(LEFT(L2#,3),Clients!A:A,Clients!B:B)
I2: =ROUND(SUMIFS(Raw_ALL!AE:AE,Raw_ALL!AF:AF,L2#),0)
J2: =IF(ISNUMBER(SEARCH("|HDC|",L2#)),SUMIFS(Raw_ALL!I:I,Raw_ALL!AF:AF,L2#),"")
```

**What they do:**

- **L2:** the list of distinct keys, one per submission row. FILTER drops excluded rows, UNIQUE keeps one of each, SORT tidies.
- **B2:** joins all keys into one long text with ";" between them, then splits it back: ";" starts a new row and "|" a new column. One formula fills a 446 x 7 grid.
- **A2:** the first 3 letters of each key are the client code; look up its organisation_id.
- **I2:** for each key, add up final_kg for all matching rows, rounded to whole kg as the format requires.
- **J2:** units only on HDC rows (the key contains "|HDC|"), blank otherwise.

**Why L2#:** the `#` means "the whole spilled list starting in L2". Every formula spills by itself, so there is no fill-down and no chance of skipping rows.

### Submission checks

```
=ROWS(L2#)
=SUM(Raw_ALL!AE:AE)
=SUM(I2#)
=SUM(--ISNUMBER(SEARCH("||",L2#)))
=COUNTIF(I2#,"<1")
```

**What they do:** row count; both totals (should be within rounding of each other); keys with an empty part (should be 0); weights under 1 kg (should be 0).

**Why the double minus (--):** it turns TRUE/FALSE into 1/0 so they can be added up.

### Debugging the zero weights

```
=Raw_ALL!AE1
=COUNTIF(Raw_ALL!AF:AF,L2)
=ISNUMBER(Raw_ALL!AE2)
```

**What they do:** check the column is the right one, check the key matches, check the value is a number.

**Why:** three separate questions, each ruling one cause in or out. The third one showed final_kg was empty.

### Marking against the answer key

```
L2: =B2:B447&"|"&C2:C447&"|"&D2:D447&"|"&E2:E447&"|"&F2:F447&"|"&G2:G447&"|"&H2:H447
M2: =XLOOKUP(L2#,Submission!L2#,Submission!I2#,"MISSING")
    =COUNTIF(M2#,"MISSING")
N2: =M2#-I2:I447
    =SUM(--(ABS(N2#)>1))
O2: =XLOOKUP(L2#,Submission!L2#,Submission!J2#,"MISSING")
    =SUM(--(O2#&""<>J2:J447&""))
```

**What they do:** build the same key on the answer key, bring back my weight and units for each row, and count missing rows, weight differences over 1 kg, and unit differences.

**Why the &"" part:** it turns both sides into text, so a blank compares cleanly with a blank and a number with a number.

**Result:** 8 missing (the MDC bug), then 0, 0 and 0 after the fix.

## F. Swings and benchmark (Module 7)

### Swings sheet

```
A2: =SORT(UNIQUE(INDEX(Submission!B2#,,1)&"|"&INDEX(Submission!B2#,,6)))
B2: =SUMIFS(Submission!I:I,Submission!B:B,LEFT(A2#,3),Submission!G:G,RIGHT(A2#,2))
C2: =SUMIFS(Prior_2024!I:I,Prior_2024!B:B,LEFT(A2#,3),Prior_2024!G:G,RIGHT(A2#,2))
D2: =IF(C2#=0,"NEW",B2#/C2#-1)
E2: =IF(ISNUMBER(D2#),IF(ABS(D2#)>0.2,"SWING",""),"NEW")
    =COUNTIF(E2#,"SWING")
```

**What they do:**

- **A2:** `INDEX(spill,,1)` takes column 1 (client) and `INDEX(spill,,6)` takes column 6 (material) of the Submission grid, giving keys like `AVD|GL`.
- **B2, C2:** total kg for 2025 and 2024. The client code is always 3 letters (LEFT) and the material always 2 (RIGHT).
- **D2:** % change. If there was nothing last year, show NEW instead of a divide-by-zero error.
- **E2:** flag any move bigger than 20% in either direction (ABS).

**Why client and material level:** that is how account managers talk about a client ("AVD glass"). Nation and class are too detailed for a first look.

### Benchmark

```
G2: =B2#/SUMIFS(Submission!I:I,Submission!B:B,LEFT(A2#,3))
H2: =XLOOKUP(RIGHT(A2#,2),Benchmark!A:A,Benchmark!B:B)
I2: =(G2#-H2#)*100
```

**What they do:** the material's share of the client's total, the UK share for the same material, and the gap in percentage points.

**Why percentage points:** a share going from 1% to 2% is "+100%" but only "+1 point". Points give an honest sense of size.

## G. Python used in Google Colab

### GGS: flattening the JSON

```python
from google.colab import files
import json, pandas as pd

uploaded = files.upload()
with open('GGS_GloucesterGarden_ERP_export_2025.json') as f:
    data = json.load(f)

df = pd.json_normalize(data['records'])
print(df['ship_to_region'].value_counts())
df.to_csv('GGS_records.csv', index=False)
files.download('GGS_records.csv')
```

**Why:** web Excel had no Power Query to open JSON. `json_normalize` turns nested records into a flat table, and `value_counts` showed the 5 regions, including Ireland.

### CBH: reading a European CSV

```python
df = pd.read_csv('CBH_CardiffBayHomewares_2025.csv',
                 sep=';', encoding='cp1252',
                 decimal=',', thousands='.',
                 dtype={'Period_end': str})
df = df[df['SKU'].notna()]   # drop the client's TOTAL row
```

**Why each setting:** semicolon separator; Windows-1252 encoding so Café and Crème read correctly; comma as the decimal point; dot as the thousands separator. The date column is read as text, because the "dot is a thousands separator" rule would otherwise turn 30.06.2025 into a number.

### CCR: joining sales to specs

```python
sales = pd.read_excel('CCR_CliftonCoffee_sales_2025.xlsx')
specs = pd.read_csv('CCR_CliftonCoffee_packaging_specs.csv')

m = sales.merge(specs, on='SKU', how='left')

long = m.melt(id_vars=['SKU','Product','Channel','Nation','Component',
                       'Material','Class','Unit weight (g)','Spec date'],
              value_vars=['H1 Units','H2 Units'],
              var_name='raw_period', value_name='units')

long['raw_weight'] = long['units'] * long['Unit weight (g)']

print(long[long['Component'].isna()]['SKU'].unique())
print(long.groupby('SKU')['Spec date'].nunique().loc[lambda s: s > 1])

long.to_excel('CCR_joined_long.xlsx', index=False)
```

**Why each step:**

- `how='left'` keeps every sale, even with no spec, so missing specs show up as blanks to query instead of disappearing.
- `melt` unpivots H1 Units and H2 Units into one column with a period label.
- The two `print` lines are checks: SKUs with no spec (CCR-0038, 0044, 0045), and SKUs with more than one spec date (CCR-0002).

**Why Python here:** one sales row matches several spec rows (many-to-many). An Excel lookup returns only the first match and would silently get it wrong.
