# The full story

## 1. Why I started

On Monday 21 September 2026 I had my first-stage phone screen with Valpak for the **EPR Data Consultant (US)** role. It went well, but one gap came up clearly: they asked about complex Excel and **VBA**, and VBA was the weakest part of my answer. A few days later I passed the Thrive assessment (Numerical 91, Verbal 90), and then there was nothing to do but wait for a possible second interview.

I was on a 1:36am bus to a Starbucks shift when I decided I did not want to just wait. I wanted to be able to say:

> "Talk is cheap. Here is my project."

The plan was to build a project that does what a Valpak data consultant actually does: take messy packaging data from clients, clean it, check it, query the problems, build the final submission, look for swings, and automate the repetitive parts with VBA. Closing the VBA gap on my own initiative would be the story.

### Keeping the scope honest

My first thought was to go all the way to "data-driven decision making" with dashboards and predictions. I decided against that. The job description talks about collating, manipulating, checking and submitting data, swings analysis and trends, and finding efficiencies. So the project matches the JD and nothing more. A focused project that mirrors the real job is more convincing than a big one that does not.

## 2. Setting up the project

### Where the data came from

I wanted UK data because familiar places make it feel real, even though the role is US-focused. The real source is the Environment Agency's **National Packaging Waste Database (NPWD)**, which publishes the pEPR Reported Packaging Data. I downloaded the file `pEPR_2023_2024_and_2025_Reported_Packaging_Data_09092026_for_publishing.xlsx`.

Real supplier-level data is never published, so the project uses a mix:

- **Real** UK EPR categories and codes, and real national totals (the "Total Supplied" row from the LP_2024 and LP_2025 weight summary sheets) for the benchmark.
- **Fictional** clients and messy files built on top, so that every typical problem appears at least once. The practice pack was generated for me (with a fixed random seed) so that there was a hidden answer key I could mark myself against later.

### The UK packaging EPR codes

These codes were the "language" of the whole project:

| Field | Codes |
|---|---|
| Material | AL aluminium, FC fibre composite, GL glass, OT other, PC paper/card, PL plastic, ST steel, WD wood |
| Packaging type | HH household, NH non-household, HDC household drinks container |
| Packaging class | P1 primary, P2 secondary, P3 shipment, P4 tertiary |
| Nation | EN England, SC Scotland, WS Wales, NI Northern Ireland |
| Activity | SO brand owner, PF packer/filler, IM importer |

Two important rules: **HDC is only allowed for AL, GL, PL and ST**, and HDC rows **must have units** (number of containers). The Republic of Ireland is **not** part of UK EPR.

### The practice pack

| Folder | What is in it |
|---|---|
| 01_client_submissions | The raw client files. Never edit these. |
| 02_reference | client_register.xlsx, target_codes_and_format.xlsx, EPIC_prior_submissions_2024.csv, national_benchmark_UK.xlsx |
| 03_client_replies | Reply emails from clients, plus a late SPF file. Only opened after sending queries. |
| 04_your_work | Everything I build |
| _answer_key | expected_final_submission_2025.csv (446 rows) and issues_log.csv (71 issues). Not opened until Module 6. |

The target submission format has 10 columns: organisation_id, client_id, submission_period, packaging_activity, packaging_type, packaging_class, packaging_material, nation, packaging_material_weight_kg, and packaging_material_units (HDC rows only). One row per client, period, activity, type, class, material and nation, with weights summed.

### The 8 clients and what was wrong with each file

| Client | File | What made it messy |
|---|---|---|
| AVD Avon Valley Drinks | xlsx, close to the template | Codes mixed with full names ("Plastic", "PL", "Plastic  "), the nation "Nor", several period spellings (H1 2025, 2025 H1, Jan-Jun 25), about 3% trailing spaces, 12 blank weights, 4 weights typed in tonnes in a kg column, 6 exact duplicate rows, HDC rows needing units, and a real glass-to-cans change vs 2024 |
| SSN Severnside Snacks | CSV export from an ERP system | Free-text labels and typos (Plastik, corrugated, prim, B2B, N. Ireland), weights in tonnes with a trailing "t", 8 "n/a" weights, 5 negative stock-adjustment lines, a page-break line and repeated header every 500 rows |
| BBO Bath Botanicals | xlsx report layout | One sheet per nation, title rows, merged H1/H2 headers, subtotal and total rows, numbers stored as text with commas, no packaging type, 5 "TBC" weights, free-text materials (Glass (jar), Alu tube, PP plastic, HDPE, Bamboo, Cotton) |
| CBH Cardiff Bay Homewares | CSV, European format | Windows-1252 encoding (Café, Crème), semicolon separator, decimal comma, dot as thousands separator (12.500 means 12,500), weights in grams, its own codes (PAP, PLA, MET-ST, WOO; PRI, SHP, TER; ENG, WAL), a TOTAL row, no packaging type |
| MDC Mendip Dairy | xlsx summary | Weights in tonnes, no nation (register says England only), type and class merged into one label, milk cartons reported as "HH Drinks, Paper / Card" (an invalid HDC), and one typed row total wrong by 41.5 tonnes |
| CCR Clifton Coffee Roasters | Sales xlsx plus a specs csv | Sales units in one file and per-unit packaging weights in another, so they must be joined. 3 SKUs sold with no spec, two spec versions for CCR-0002, the sales channel decides HH or NH, and real growth vs 2024 |
| GGS Gloucester Garden Supplies | JSON from a US parent system | Weights in pounds, US material codes (LDPE, PP, HDPE, CORRUGATED, PAPERBOARD, KRAFT), MM/DD/YYYY dates, 132 rows shipped to the Republic of Ireland |
| SPF Swindon Pet Foods | xlsx, clean looking | The whole H2 period missing, and one hidden x10 decimal slip (SPF-0070 Paper sack) |

### The module plan

| Module | What | Status |
|---|---|---|
| 0 | Know the rules (EPR codes and format) | Done |
| 1 | Intake and completeness check | Done |
| 2 | One table for everything (Raw_ALL) | Done |
| 3 | Map and standardise | Done |
| 4 | Validation rules | Done |
| 5 | Queries, client replies, overrides | Done |
| 6 | Build the final submission and mark it | Done, 100% match |
| 7 | Swings analysis and benchmark | Done |
| 8 | Automate with VBA | Done |
| 9 | US version (stretch) | Optional |

VBA was deliberately left to the end. The idea was: **do each step by hand first, understand it, then automate it.** You cannot automate a process you do not understand.

### My tools

I started with **Excel for the web only** (no desktop Excel, so no VBA and no Power Query) and used **Google Colab (Python)** only where web Excel struggled. Later I bought Microsoft 365 Personal for desktop Excel so I could do the VBA module.

## 3. Module 1: Intake, looking before touching

The first job was not to fix anything. It was to **open every file and write down what I saw**, in an intake log with 10 columns: client, file names, file type, layout, periods, nations, weight unit, target fields missing, first-look problems, and ready to process.

My first draft had weak entries. Feedback I got:

- "No organisation_id" is not a finding. That comes from the client register, not from the client.
- Only mark a file "not ready" if it is truly blocked.
- For BBO, the nation is hidden in the sheet names.
- For CBH, "difficult to identify" is not acceptable. Open the file in Notepad and look at the raw text.
- For CCR, count the distinct SKUs, not the rows.

My second draft got the numbers right: AVD 12 blank weights, SSN 52 values with a trailing "t", 8 "n/a", 2 repeated headers and 5 negatives. For CCR I first wrote 142 spec rows vs 135 sales rows. Those were row counts. The real comparison is **42 distinct SKUs with specs vs 45 SKUs sold**, which is how I found the three SKUs with no spec (CCR-0038, 0044 and 0045). GGS had 5 regions, one of them Ireland.

My best catch was SPF: the file only had H1, so I flagged that H2 was missing and needed chasing.

### Python where Excel struggled

Web Excel had no Power Query, so I used Colab for two files:

- **GGS (JSON):** I loaded the JSON, flattened it with `pd.json_normalize`, counted the regions, and saved a CSV.
- **CBH (European CSV):** I read it with a semicolon separator and the right encoding. For the Qty column I removed the "." thousands separators in Excel instead.

One small moment taught me something. I asked whether to remove the commas in CBH's weights too. No: in CBH, **the comma is the decimal point**, so it has to become a dot. Removing it would make every weight 10 to 1,000 times too big. The same character means different things in different files.

### Excel vs Python

I asked whether Valpak would care more about Excel than Python. The decision: **Excel for the process, Python only where Excel genuinely struggles** (the JSON, the CBH separators, and the CCR many-to-many join). The interview line is: "I used Excel for the pipeline, and Python for one join that a lookup would have got wrong."

## 4. Module 2: Building one table (Raw_ALL)

### The CCR join, my first real concept

CCR had sales in one file (units per SKU, nation and half-year) and packaging specs in another (grams per component per SKU). To get weights I had to join them.

At first I mixed up the two files. I thought CCR-0002 had 3 spec rows and 8 sales rows. It is the other way round: **8 spec rows** (4 components x 2 spec versions) and **3 sales rows** (one per nation). Getting the **grain** of each file right (what one row means) was the first lesson.

The other lessons from this join:

- **Duplicate spec versions double the weight.** CCR-0002 had a 2024 spec and a 2025 spec. Joining 3 sales rows to 8 spec rows gives 24 rows, with every component counted twice.
- **Left join, not inner join.** An inner join silently drops the 3 SKUs with no spec. For CCR-0044 alone that is about 38,000 bags of coffee that would just disappear. A left join keeps them with blank weights, so they show up as a problem to query.
- **Unpivot after the join.** H1 Units and H2 Units become one "units" column with a period label.
- **weight = units x grams per unit.**

I was exhausted that night (a full day of work plus a driving lesson) and told myself "I can only understand half". I asked to see it with **real rows** instead of a toy example, because I never run code I don't understand. Seeing CCR-0002 England H1 (10,373 units x 9.5 g valve bag = 98,543.5 g) made it click.

### Loading every client into one table

Raw_ALL has one row per data line from every client, with the values **exactly as the client sent them**: client_id, sku, component, raw_material, raw_class, raw_type, raw_nation, raw_period, units, raw_weight, raw_weight_unit, note. Junk lines (headers, page breaks, subtotals, the CBH TOTAL row) were skipped, but suspicious data (negatives, n/a, Ireland) was kept, because those are things to query, not to hide.

Expected rows per client: AVD 1,068, SSN 1,455, CBH 1,920, GGS 1,782, SPF 752, CCR 870, BBO 2,383, MDC 18. Total 10,248.

### Problem: the row count was wrong, twice

My first count was **10,266**, which was 36 too many. Then a recount said **20,135**.

How I found it: a **per-client reconciliation table**, with `=COUNTIF(Raw_All!A:A,"AVD")` for each client next to the expected count and a difference column. That showed everything was exact except **BBO at 2,388 (+5)**. The 20,135 meant about 9,900 rows had no client_id, from pastes that did not line up.

- The extra 5 BBO rows were product lines with an empty H1 or H2 block. Fix: filter units blank AND weight blank, then delete.
- The unlabelled rows: Ctrl+End to find the real end, filter client_id = blanks, delete.

Result: exactly 10,230, then 10,248 with MDC added.

**Lesson: always reconcile counts per group, not just the total.** A total can hide two errors that cancel out.

## 5. Module 3: Mapping, turning client language into codes

### Listing what the clients actually wrote

I listed every distinct value in each raw column with `=UNIQUE(...)`. Things I learned from those lists:

- A "0" in a UNIQUE result is a **blank** cell.
- UNIQUE and XLOOKUP ignore case, so "PRIMARY" and "Primary" fold together.
- 45838 and 46022 are Excel date serial numbers for 30 Jun 2025 and 31 Dec 2025.

Some mappings needed real judgement:

- Composite, Paper tube (composite) and Fibre composite go to **FC**, not PC.
- Tinplate goes to **ST**. Alu, Alu tube and Aluminium foil go to **AL**.
- Cotton and Bamboo go to **OT** (other).
- PLA goes to **PL** (CBH's code for plastic, and also a bioplastic).
- Outer, Transit and SHIPPING go to **P3**. Pallet goes to **P4**.
- Retail and Online shop mean **HH**. Wholesale and Cafe mean **NH**.
- IRL goes to **OUT_OF_SCOPE**, so the exclusion is deliberate and visible.

### The Mappings table

At first I did not understand why I needed a Mappings table at all. For plastic alone there were Plastic, PL, "Plastic  ", "PL  ", PP plastic, Plastik, plastic film, Plastics, PLA and PP. Why not just Find & Replace?

The answer: the Mappings table is a **dictionary**. Many spellings point to one code. It is better than Find & Replace because:

- It **keeps the raw value**, so I can always see what the client actually wrote, for example when writing a query.
- It is **reusable**. Next quarter, the same table maps the new files.
- It is **checkable**. Anyone can review the judgement calls in one place.

The table has columns field, raw_value, target, judgement, note, and a key column `=A2&"|"&B2`, which gives keys like `material|Plastik`. The **field** part matters because the same text can mean different things in different columns: "HH Drinks - Primary" means **HDC** as a type but **P1** as a class.

Each clean column then looks the key up, for example:

`=IF(D2="","",XLOOKUP("material|"&TRIM(D2),Mappings!F:F,Mappings!C:C,"UNMAPPED"))`

TRIM handles the trailing spaces, and the "UNMAPPED" default makes any spelling I forgot show up loudly instead of silently becoming blank.

I also asked how to replace raw_material with clean_material. The answer: **don't**. Keep both columns side by side. That way a mapping change updates everything, queries can quote what the client wrote, and there is an audit trail.

### Rules that a lookup alone cannot handle

Some fields needed logic, written down on a Rules sheet:

1. **MDC nation:** MDC gave no nation, but the client register says England only, so EN.
2. **MDC class:** MDC's class is hidden inside its type label ("NHH - Shipment" means NH and P3), so for MDC the class formula looks up the type label instead.
3. **BBO and CBH packaging type:** neither gave a type, so it is assumed from the class (P1/P2 means HH, P3/P4 means NH) and marked "Type assumed: confirm with client".
4. **CCR shipping cases:** a shipping case (P3) is always NH, even if the channel is "Retail". This check has to come **first** in the formula, otherwise "Retail" wins.

This part was hard. At one point I wrote "I don't know where to start" and later "just give me the answer". What finally worked was going through each rule with real rows, like the CCR-0001 shipping case on Retail showing HH when it should be NH.

### Converting weights to kg

The weight_kg formula converts every unit to kg (tonnes x1000, g x0.001, lbs divided by 2.20462, not 2.2) and cleans the text at the same time: it strips SSN's trailing "t", removes BBO's thousands commas, and turns CBH's decimal comma into a dot. Anything that is not a number (n/a, TBC) stays blank, because those are **queries, not zeros**.

### Problem: Find & Replace stripped every comma from my formulas

I selected the raw_weight column and used Find & Replace to remove commas. In Excel for the web, the replace was not limited to my selection. It **removed every comma from every formula** in the sheet, so all my XLOOKUPs broke ("TRIM(G5) Mappings!$F$2:$F$122 Mappings!..."). I wrote: "something happened, help."

Fix: **version history.** I restored the version from before the replace. Then, instead of ever using Find & Replace on the data again, I moved the comma handling **inside the weight_kg formula**, where it depends on the unit: for grams (CBH) the comma becomes a dot, for everything else it is removed.

**Lesson: never bulk-edit raw data. Put the cleaning logic in a formula, where it is visible and reversible.**

### Problem: every row showed the same weight

Right after that, weight_kg showed 23 on every row. Every formula pointed at J2 and K2. I had pasted the formula **text** into every cell instead of filling it down, so the row references never moved.

Fix: clear the column below row 2, then fill down from S2 (fill handle or Ctrl+D), and check that S3 refers to J3 and K3.

## 6. Module 4: Validation, letting the data tell me what is wrong

I added flag columns, each one testing one kind of problem:

| Flag | Count | Meaning |
|---|---|---|
| MISSING WEIGHT | 43 | AVD 12, SSN 8, BBO 5, CCR 18 (the 3 SKUs with no spec) |
| NEGATIVE OR ZERO | 5 | SSN stock adjustments |
| UNMAPPED | 0 | Every spelling mapped |
| OUT OF SCOPE | 132 | GGS rows for the Republic of Ireland |
| INVALID HDC MATERIAL | 2 | MDC milk cartons reported as paper drinks containers |
| DUPLICATE | 12 | 6 true AVD pairs (after fixing the key, see below) |

### Problem: false duplicates

My first duplicate check flagged rows that were not duplicates, for example two BBO lines with 3,252 and 2,320 units. My duplicate key joined client, SKU, component, material, type, nation, period, weight and note, but **not units**, so two different lines with the same weight looked identical.

I added units to the key, but it still did not work. The reason: I had **filled the new formula down with a filter on**, and Excel only fills the **visible** rows. The hidden rows kept the old formula.

Fix: clear the filters, fill down again. Result: 12 rows, all AVD, which are exactly the 6 real duplicate pairs.

**Lesson: never fill down with a filter on.**

### The weight-per-unit check

Some errors pass every other check. The value is positive, mapped and unique, but still wrong. For example, if 99,386 glass bottles weigh 18.9 kg, that is 0.19 g per bottle, which is impossible. The real figure is about 190 g.

So for every row with units I calculated **grams per unit**, and compared it with the **median** grams per unit for the same SKU and component. A ratio near 0.001 means tonnes were typed as kg. A ratio near 10 means an extra zero.

**Why median and not average:** I tested both. In a group with one x10 error, the wrong value drags the average up so far that the three **correct** rows get flagged and the wrong one does not. The median ignores one extreme value.

To keep web Excel fast, the median for each of the roughly 1,770 SKU and component groups was calculated once on a separate UnitCheck sheet, not on every row.

Result: **exactly 5 rows flagged.**

| Row | Ratio | Meaning |
|---|---|---|
| AVD-0009 Glass bottle | 0.0010 | Tonnes typed as kg |
| AVD-0003 4-pack carton | 0.0010 | Tonnes typed as kg |
| AVD-0020 PET bottle | 0.0010 | Tonnes typed as kg |
| AVD-0029 One-trip wooden pallet | 0.0010 | Tonnes typed as kg |
| SPF-0070 Paper sack | 10.0 | Extra zero (404,620 kg instead of 40,462 kg) |

## 7. Module 5: Queries, replies and corrections

### Asking the clients

Every flag that I could not fix with certainty became a query. To save time on a busy week, I asked for the query and reply scenario to be run for me, and then read through it in detail: the query log, the 8 emails, the replies and the list of changes to apply. In real work the consultant writes these emails; the important skill here is knowing **what to ask, what to decide yourself, and how to apply the answers**.

The query log had **20 items**: 16 closed, 2 still open and 2 parked for the swings analysis.

| ID | Issue | Outcome |
|---|---|---|
| AVD-01 | 12 missing weights | Weights supplied |
| AVD-02 | 4 weights 1,000x too small | Confirmed tonnes, multiply by 1,000 |
| AVD-03 | 6 duplicate pairs | Decided to remove one of each, client told |
| AVD-04 | "Nor" read as NI | No answer, still open |
| AVD-05 | Glass down, cans up | Parked for Module 7 |
| SSN-01 | 8 "n/a" weights | Weights supplied |
| SSN-02 | 5 negative lines | ERP stock adjustments, exclude |
| BBO-01 | 5 "TBC" weights | Weights supplied |
| BBO-02 | Assumed packaging type | Confirmed |
| CBH-01 | How the file was read | Confirmed grams, dot as thousands |
| CBH-02 | Assumed packaging type | Reply answered a different question, still open |
| MDC-01 | Paper/card as HDC | They are fibre composite milk cartons, household primary |
| MDC-02 | Typed H2 total wrong | The detail is right |
| MDC-03 | Nation | Confirmed England only |
| CCR-01 | 3 SKUs with no spec | 11 spec lines supplied (new 2025 products) |
| CCR-02 | Two spec versions | Use the 2025 spec, the bag was made lighter in January 2025 |
| CCR-03 | Growth vs 2024 | Parked for Module 7 |
| GGS-01 | 132 Ireland rows | Dublin stockists, leave out |
| SPF-01 | H2 missing | Late H2 file sent (752 rows) |
| SPF-02 | x10 slip | Confirmed typo |

Lessons from the replies: a reply can look complete and not be (CBH-02), clients often explain changes without being asked, and "excluded" is not the same as "deleted".

### Applying the answers without destroying the evidence

The principle I followed: **never overwrite the raw value**. Instead I added three columns:

- **override_kg:** the corrected weight from the client, typed in. 30 cells in total: 25 missing weights (AVD 12, SSN 8, BBO 5) plus the 4 AVD tonnes fixes and SPF-0070.
- **exclude_reason:** a formula that gives a reason when a row should not be in the submission: duplicates (only the second copy of each pair), Ireland, stock adjustments, and later the old CCR rows.
- **final_kg:** blank if excluded, otherwise the override if there is one, otherwise the converted weight.

For MDC's milk cartons, I set clean_material to FC and clean_type to HH on those 2 rows by hand, with a note "Per client email MDC-01". After that, INVALID HDC MATERIAL dropped to 0.

The MISSING WEIGHT flag still says 43 after all this, and that is correct. **Flags describe what the client sent. final_kg holds the correction.** Both stay visible.

### Problem: I overwrote 26 SKUs while typing corrections

While typing the missing weights with the MISSING WEIGHT filter on, I accidentally **filled "AVD-0002" down over all the visible rows**.

How I found it: `=COUNTIF(B:B,"AVD-0002")` returned **72**, but the original file had **46**. So 26 rows had the wrong SKU: 10 AVD, 3 BBO, 8 SSN and 5 CCR. I wrote: "there are 72, what happened to them, omg."

I did not want to roll back, because I would lose the columns I had added that day. Instead the 26 correct SKUs were looked up in the source files by matching component, nation, period and units, and I retyped them **one cell at a time**. Later, my SKUs matched the client's reply exactly, which confirmed the repair.

Checks to prove it was clean:

- `=COUNTIFS(A:A,"<>AVD",B:B,"AVD*")` should be 0 (no AVD SKU on a non-AVD row).
- A SUMPRODUCT that counts rows where the first 3 letters of the SKU do not match the client should also be 0.

**Lesson: with a filter on, type one cell at a time. Never fill or paste over a filtered range.**

### Problem: 4 real BBO rows had disappeared

The status bar said 10,244 records and BBO showed 2,379 instead of 2,383. Back in Module 2, when I deleted the 5 empty BBO lines, some deletions landed **one row off**, so I deleted real rows next to the empty ones. The totals had still looked right.

How I found exactly which rows: a file of the 2,383 expected BBO rows, and an **anti-join** (find rows in the expected list that are not in Raw_ALL):

`=COUNTIFS(Raw_ALL!B:B,B2,Raw_ALL!C:C,C2,Raw_ALL!P:P,M2,Raw_ALL!Q:Q,N2)`

Exactly 4 rows returned 0: BBO-0090 Shipping case (SC H1), BBO-0002 Carton (SC H1), BBO-0007 Shipping case (SC H1) and BBO-0090 Bamboo box (WS H2). I added them back.

**Lesson: delete through a filter, never by eye.**

### Problem: adding those 4 rows overwrote 3 CBH rows

After pasting the 4 BBO rows, CBH dropped to 1,917. I had pasted them over the first rows of the CBH block. I had assumed Excel would make room by itself. It doesn't: **paste always overwrites**. I fixed it the same way, with an expected-rows file for CBH and the same COUNTIFS check, and from then on always pasted below Ctrl+End.

As I wrote at the time: "now I know lol, when I paste BBO, I didn't insert 4 new rows, I thought it would create itself."

### Problem: "the filter shows 43 missing weights, but the replies only have 37"

This confused me. The answer: 25 of the 43 are typed in as overrides (AVD 12, SSN 8, BBO 5), and the other 18 are the CCR placeholder rows, which are not typed at all. They get rebuilt by re-joining CCR with the 11 new spec lines.

I also noticed the client listed the SKUs in a different order from mine. That does not matter: **always match on SKU + component + nation + period, never on position.**

### Rebuilding CCR

With the 11 new spec lines and the instruction to use the 2025 spec, I rebuilt CCR: dropped the 4 old CCR-0002 spec lines, added the 11 new ones, left-joined, and unpivoted. That gave **894 rows**, marked "rejoin (CCR-01/02)". I pasted them at the bottom and the old 870 CCR rows were excluded by formula with the reason "Replaced by CCR re-join". Nothing was deleted.

### Problem: 752 UNMAPPED after adding SPF's late file

When I added SPF's 752-row H2 file, the UNMAPPED count jumped to **exactly 752**. The number matching the new rows exactly told me the problem was something the whole new file had in common.

Diagnosis, step by step:

- `=LEN(Raw_ALL!H11895)` returned **7**, so the period text was exactly "2025-H2" with no hidden spaces.
- `=COUNTIF(Mappings!F:F,"period|2025-H2")` returned **0**, so the Mappings table had no row for that spelling.

When I pasted my H2 mapping rows, there were H2 2025, 2025 H2, Jul-Dec 25 and others, but **no "2025-H2"**. I had never added it. Also, my lookup used a fixed range ($F$2:$F$122), so a new mapping row at row 123 would have been ignored anyway.

Fix: add the period|2025-H2 row, and switch clean_period to **whole columns** (Mappings!F:F and Mappings!C:C) so new mapping rows are always included.

### End of Module 5

| | Rows |
|---|---|
| Rows in Raw_ALL | 11,894 |
| Excluded, each with a written reason | 1,013 (870 old CCR, 132 GGS Ireland, 5 SSN stock adjustments, 6 AVD duplicates) |
| Included in the submission | 10,881 |

I studied for 4 to 5 hours that day and finished with "I still don't understand, lol, should I rest now?" I rested. Module 5 was finished on Friday 2 October.

## 8. Module 6: Building the submission and marking it

### Squeezing 10,881 rows into a few hundred

The submission needs one row per client, period, activity, type, class, material and nation. I did it with the same "build a key, then use it" pattern as the Mappings table, except this time to **add up** instead of to look up.

1. A small **Clients** sheet with each client's organisation_id and activity (SO, PF or IM).
2. A **sub_key** column on Raw_ALL that glues the 7 fields together, for example `AVD|2025-H1|SO|HDC|P1|AL|EN`. Excluded rows get a blank key, so they drop out on their own.
3. A **Submission** sheet built from 5 spilling formulas: UNIQUE for the list of keys, TEXTSPLIT to split each key back into 7 columns, XLOOKUP for organisation_id, SUMIFS for the weight, and a second SUMIFS for units on HDC rows only.

Using spilling formulas (with `L2#` meaning "the whole list that starts in L2") meant **no fill-down at all**, which removed the cause of several earlier accidents.

My first key list had **440** rows.

### Problem: every weight was 0

The first weight came out as 0. I checked step by step: AE1 was the final_kg header (correct column), the first key matched 8 rows on Raw_ALL (keys were fine), but `=ISNUMBER(Raw_ALL!AE2)` was FALSE. When I pasted two rows, final_kg was empty. **The final_kg column had no formula at all.** It had either never been filled down or something had wiped it.

Fix: re-enter `=IF(AD2<>"","",IF(AC2<>"",AC2,S2))` and fill it down. The first submission row then showed **13,709 kg**, and its units showed **945,349**, both matching the answer key exactly.

### Sanity checks before marking

- Total kg on Submission: **9,131,217**. Total kg on Raw_ALL: **9,131,216**. Only 1 kg apart, from rounding. Nothing lost, nothing doubled.
- No key contains "||" (an empty part, such as a blank class).

### Marking against the answer key

I pasted the answer key into its own sheet, built the same key on it, and used XLOOKUP to bring back my weight for each of its rows. The answer key had **446** rows and I had **440**. **8 of its keys were MISSING** from mine, which meant I also had **2 extra keys**.

### Problem: the case of the MDC rows that pointed at row 5

This was the best piece of tracing in the whole project.

1. **The totals matched exactly, but the row counts did not.** So no weight was lost. Rows had been **labelled differently** and merged.
2. **All 8 missing keys were MDC, and all were class P2, P3 or P4.** My MDC rows only had P1. So the problem was in the class column.
3. If Secondary, Shipment and Tertiary had all become P1, then HH P2 PC would become a new key HH P1 PC (my 2 extra keys, one per half-year), and the others would merge into existing P1 keys. That explained 8 missing and 2 extra exactly.
4. The Mappings table returned **P2** for "HH - Secondary", so the dictionary was fine.
5. The formula on the MDC rows was the right formula, but it referred to **E5, F5 and A5**. Every MDC row (from about row 8,025) was reading **row 5**, which is an AVD row. AVD's raw_class mapped to P1, so every MDC row became P1.

I also found that column N had **different versions** of the formula in different client blocks. The GGS rows above had an older version without the MDC part.

Fix: type the correct formula on the first MDC row, with its own row number and whole-column Mappings ranges, and fill down over the MDC rows only. MISSING dropped to 0.

### The result

- **446 rows, 0 missing.**
- **0 rows** with a weight difference of more than 1 kg.
- **0 rows** with a units difference.

**My submission matched the answer key 100%.**

### Reviewing the issues log

The answer key came with an issues log of **71 planted issues**. My numbers matched, so I had dealt with all of them, but two I had not consciously noticed:

- **AVD trailing spaces:** handled anyway, because every lookup has TRIM inside it.
- **MDC's typed total:** the TOTAL cell for H2 household primary said 516.718 tonnes, but the materials summed to 475.218. I avoided it only because I never used the TOTAL column. The real lesson: **a client's typed total is a free check. Recalculate it, and if it doesn't match, ask which is right.**

## 9. Module 7: Swings and the benchmark

### Swings: comparing each client with last year

I built a **Swings** sheet with one row per client and material (38 rows), the 2025 and 2024 totals (both with SUMIFS), the % change, and a flag for anything that moved more than 20% in either direction.

**7 swings** were flagged: AVD aluminium +31%, and all 6 of CCR's materials, up by roughly 90% to 115%.

### My first reaction was wrong

I said they were all errors, and that 35,230 kg vs 26,898 kg could not be a 30% difference. Both parts of that were wrong:

- **The maths:** 35,230 - 26,898 = 8,332 kg more, and 8,332 / 26,898 = 0.31, so a **31% rise**. Percentage change is measured against last year's figure.
- **The judgement:** a swing is a **question**, not a verdict. Both clients had already explained in Module 5. AVD moved several beers from glass bottles to cans, and CCR took on new cafe and wholesale accounts. Both were real business changes, so the data stays and the reason is written down.

How to tell real from error, beyond the email:

- **AVD:** aluminium up **and** glass down at the same time. That is one consistent story: bottles became cans.
- **CCR:** **every** material roughly doubled. That is what business growth looks like.
- **A typical error** hits **one** item, like SPF-0070's x10 slip, which affected one product in one nation in one half-year.

### The one the flag missed

AVD glass was **-19.3%**, just under the 20% line, so it was not flagged. But it is the other half of the same bottles-to-cans story. **A threshold is a tool, not a rule.** A good analyst also reads the rows just under the line, especially when they connect to a swing that was flagged.

### The benchmark: comparing each client with the UK

A swing compares a client with itself. A benchmark compares a client with the whole UK. I calculated each material's share of each client's total and compared it with the UK share from the national data, measured in **percentage points**. (2% vs 1% is "+100%", which sounds dramatic but means nothing. "+1 point" is honest.)

The biggest gaps were all paper, and all made sense for the business:

| Client and material | Client share | UK share | Why it makes sense |
|---|---|---|---|
| CBH paper/card | 91.7% | 42.7% | A homewares importer: goods shipped in boxes |
| SSN paper/card | 65.2% | 42.7% | Snacks: cartons and shipping cases |
| SPF paper/card | 64.1% | 42.7% | Pet food: paper sacks |
| BBO other | 6.7% | 0.2% | Bamboo and cotton gift packaging |
| MDC fibre composite | 11.7% | 1.2% | Milk cartons (the MDC-01 reclassification) |
| SPF steel | 14.8% | 3.8% | Pet food tins |

**The benchmark is a sense check, not an error finder.** The worrying case is a mix that does **not** fit the business. For example, if a dairy showed 0% fibre composite, you would ask where the milk cartons went.

## 10. Module 8: Automating it with VBA

### Getting desktop Excel

VBA does not run in Excel for the web, so I bought Microsoft 365 Personal with my coffee money. Desktop Excel was immediately easier: faster, no lag on 11,000 rows, and everything in the menus. I turned on the Developer tab and saved a macro-enabled copy (master_2025_vba.xlsm), keeping the original file as a backup.

### Recording my first macro, and what it taught me

I recorded a macro that formats the Submission sheet. Reading what the recorder wrote taught me its bad habits:

- I started recording before saving the file, so it recorded my **Save As**. Every run would have re-saved over the file. Dangerous.
- It **selects everything first** (Select, then Selection). Real VBA works on objects directly.
- It recorded noise: a scroll, an extra AutoFit on the whole sheet, and two fill colours because I changed my mind.

I rewrote it cleanly by hand.

### RunChecks: the checks I kept typing by hand, as one button

Across Modules 4 to 6 I kept typing the same COUNTIF checks into spare cells. RunChecks does them all at once: UNMAPPED codes, included rows with no kg, HDC problems, unfixed unit-weight flags, and Raw_ALL kg vs Submission kg. Then it says ALL CHECKS PASSED or SOMETHING NEEDS A LOOK. I put it on a button on the Checks sheet.

### Proving the check works

A check that has only ever said "passed" is not proven. So I broke something on purpose: I changed D2 from AL to XYZ. RunChecks reported **UNMAPPED: 1 and HDC problems: 1**. One bad cell set off two alarms, because the bad material also broke the HDC rule. The kg totals did not change at all, because the weight was fine and only the label was broken. **That is exactly why codes are checked separately from totals.**

Important VBA fact: **running a macro wipes Ctrl+Z history**, so I noted the original value first and typed it back by hand.

### Small snags

- The button first said "Cannot run the macro Button1_Click". When assigning the macro, the name box was pre-filled with a suggested new name, and I had clicked OK without picking RunChecks from the list.
- Pressing F5 in the Excel window opened the **Go To** box. F5 only runs code inside the VBA editor. From Excel, use Alt+F8.
- The export message seemed to do nothing, because the message box was hidden behind the VBA editor.

### The audit log, the export, and the gate

- **Check_Log:** every run of the checks writes one row with the time and the results. In compliance work you want to show when the checks were run and what they found.
- **ExportSubmission:** copies the submission (values only) into a new workbook and saves it as a CSV with the date and time in the name. It exported **446 rows**.
- **CheckAndExport:** the gate. It runs the checks and only exports if they all pass. I tested both ways: with D2 broken, it showed "Export stopped" and no file was created, and the log recorded a FAIL row.

### Timing it

I timed the same job both ways with a stopwatch:

| | Time |
|---|---|
| By hand: type the 5 checks, decide, copy values to a new workbook, save as CSV | about 6 minutes |
| Macro: CheckAndExport | about 3 seconds |

The macro also does things the manual run does not: it logs every run, and it refuses to export a failing file.

I exported the code as EPR_Automation.bas so I can show it without sending the workbook.

## 11. What I learned

### The habits that matter most

- **Look before you touch.** An intake log first, cleaning later.
- **Never overwrite raw data.** Keep the raw value, add a clean column next to it.
- **Exclude, don't delete.** Every excluded row keeps a written reason.
- **Put cleaning logic in formulas, not in bulk edits.** Find & Replace nearly destroyed my workbook.
- **Reconcile counts per group, not just the total.** Totals can hide errors that cancel out.
- **Never fill down or paste with a filter on.** Excel fills only the visible rows, and paste always overwrites.
- **Match rows by their key fields, never by position.**
- **Use whole-column ranges in lookups** so new rows are never missed.
- **Median, not average,** when one bad value could drag the comparison.
- **A flag is a question, not a verdict.** Swings can be real business changes.
- **Test checks in both directions.** Break something on purpose and make sure the check catches it.
- **Automate only what you understand.** Every macro in Module 8 replaced something I had already done by hand.

### How I debug

The MDC "row 5" bug showed the method I now use every time: start from what matches and what doesn't (totals matched, counts didn't), narrow it down (all MDC, all P2 to P4), check each layer separately (Mappings was fine), and only then look at the formula itself. Each step cuts the search area down.

### The final numbers

| Measure | Result |
|---|---|
| Clients | 8, in 8 different file formats |
| Rows in the clean table | 11,894 |
| Rows excluded, each with a reason | 1,013 |
| Rows in the submission | 446 |
| Match with the answer key | 100% (every row, weight and unit) |
| Client queries | 20 (16 closed, 2 open, 2 parked) |
| Swings flagged and explained | 7, plus 1 just under the threshold |
| Checks plus export | about 6 minutes by hand, about 3 seconds with VBA |

### How I would say it in an interview

> "I built a practice pipeline on real UK EPR categories: eight clients, eight different file formats, about 11,000 rows. I standardised it into one table with a mapping table, wrote validation rules that caught unit errors, duplicates and invalid categories, queried the clients, and built a submission that matched the answer key exactly. I ran a swings analysis against the prior year and a benchmark against UK totals, and then automated the checks and export in VBA, cutting about 6 minutes to 3 seconds, with an audit log and a gate that blocks a failing export. Here is what it caught, and here is what I broke along the way and how I found it."
