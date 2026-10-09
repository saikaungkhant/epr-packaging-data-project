# EPR Data Consultant practice project

You are a new EPR Data Consultant at a compliance scheme. It is February 2026. Eight clients have sent their 2025 packaging data (H1 and H2) and your job is to turn their messy files into one accurate, validated submission file, explain any big changes versus 2024, and automate the process so next year is faster.

Everything here is built on the real UK structure: material categories, packaging classes and the national totals come from the Environment Agency NPWD file you downloaded. The companies, people, emails and figures for each client are fictional.

## Folder map

| Folder | What is in it |
|---|---|
| `01_client_submissions` | The raw files, as clients sent them. Never edit these, work on copies. |
| `02_reference` | Client register, target format and code lists, last year's submitted data (your "EPIC" history), UK national benchmark. |
| `03_client_replies_OPEN_AFTER_YOU_SEND_QUERIES` | What clients say when you query them. Only open a reply after you have written the query for that client. |
| `04_your_work` | Save everything you build here. |
| `_answer_key_DO_NOT_OPEN_until_module_6` | The correct final submission and a log of every planted issue. Use it to mark yourself, not to peek. |

## The eight clients

| Code | Client | File type | Main challenge |
|---|---|---|---|
| AVD | Avon Valley Drinks, Bristol | xlsx, near template | Missing weights, duplicates, tonnes typed as kg, drinks containers need units |
| SSN | Severnside Snacks, Avonmouth | CSV export | Free text labels, typos, tonnes, n/a, negative lines, repeated headers |
| BBO | Bath Botanicals, Bath | xlsx report layout | One sheet per nation, merged headers, subtotals, numbers as text, no packaging type |
| CBH | Cardiff Bay Homewares, Cardiff | CSV, European format | Semicolons, decimal commas, grams, Windows-1252 encoding, client codes |
| MDC | Mendip Dairy, Wells | xlsx summary | Tonnes, no nation, an invalid material category, a typed total that is wrong |
| CCR | Clifton Coffee Roasters, Bristol | two files | Sales and packaging specs must be joined, missing and duplicated specs |
| GGS | Gloucester Garden Supplies, Gloucester | JSON from US parent | Pounds, US material codes, US dates, Republic of Ireland rows out of scope |
| SPF | Swindon Pet Foods, Swindon | xlsx, clean | H2 missing entirely, one hidden decimal slip |

## Modules, mapped to the job description

Work through these in order. Each module ends with something saved in `04_your_work`.

### Module 0: Know the rules (about 30 minutes)
JD: *"maintain a strong knowledge of data requirements and scoping in relevant countries"*

Read `target_codes_and_format.xlsx` until you can name the 8 material codes, 3 packaging types, 4 classes and 4 nations without looking. Read `client_register.xlsx`. Open `national_benchmark_UK.xlsx` and note which materials dominate UK packaging.

### Module 1: Intake and completeness check
JD: *"Ensure data received from customer contains necessary information"*

Open every file (just look, do not clean). Build `intake_log.xlsx` with one row per client: file type, periods present, nations present, unit of weight, which target fields are missing, first impressions. By the end you should already have spotted at least one client with a whole period missing.

### Module 2: Get everything into one table (your mini EPIC)
JD: *"Become an expert in utilising in-house software system (EPIC)... Manipulate data into relevant format"*

Load all files into one long table with the same columns: source_file, client_id, sku, component, raw_material, raw_class, raw_type, raw_nation, raw_period, units, raw_weight, raw_weight_unit. Power Query (Data > Get Data) handles CSV encoding, JSON and unpivoting. SQL or Python is fine too, use whichever you would talk about in an interview.

Hints: CBH needs encoding 1252 and a semicolon delimiter. GGS is JSON, the rows are under `records`. BBO needs unpivoting from H1/H2 columns into rows and its subtotal rows removed. CCR needs the sales file joined to the spec file on SKU.

### Module 3: Map and standardise
JD: *"Data Cleansing and Maintenance"*

Build a `Mappings` sheet: raw value in one column, target code in the next (for example `Plastik > PL`, `CORRUGATED > PC`, `N. Ireland > NI`, `Cafe > NH`). Use lookups to fill clean columns. Convert every weight to kg (tonnes x 1000, grams / 1000, lbs / 2.20462). TRIM everything.

### Module 4: Validation rules
JD: *"Use various tools to consistently check and improve on the accuracy of data"*

Add a flag column that catches at least these:
1. Any code not in the allowed lists
2. Weight blank, zero, negative or text
3. Exact duplicate rows
4. HDC rows must be AL, GL, PL or ST and must have units
5. Paper/card or fibre composite can never be HDC
6. Nation must be a UK nation
7. Both periods present for every client
8. Weight per unit (weight_kg x 1000 / units) far away from other rows for the same component. This catches unit errors and decimal slips
9. Any total typed by the client must equal the sum of its detail

### Module 5: Queries and client replies
JD: *"May involve some telephone/email support to suppliers... work alongside the customer facing account manager"*

Keep a `query_log` (client, issue, rows affected, question asked, date, answer, action taken). Draft one short, polite query email per client, grouping all their issues in one message. Only then open that client's reply in folder 03 and apply the answers. SPF's late H2 file arrives in that folder.

### Module 6: Build the final submission
JD: *"ensure data submitted to the relevant authorities complies with legal requirements"*

Aggregate your clean table to the target format: one row per client, period, activity, type, class, material and nation, with weight summed to whole kg and units for HDC rows only. SUMIFS or a Pivot Table both work.

Now mark yourself against `expected_final_submission_2025.csv`. Compare totals by client, period and material. You should be within about 1 kg per row. Then read `issues_log.csv` and tick off what you caught and what you missed. Write down why you missed each one: that list is your best interview material.

### Module 7: Swings analysis and trends
JD: *"provide useful information for accounts such as swings analysis and trends"*

Compare 2025 with `EPIC_prior_submissions_2024.csv` by client and material. Flag anything moving more than 20%. For every flag decide: data error or real business change? Some are errors, some are real, the replies tell you which. Also compare each client's material mix with the UK benchmark and note anything unusual.

### Module 8: Automate it with VBA
JD: *"Optimise data automation practices... Inspect and advance current procedures"*

Record the manual steps you repeated, then turn them into macros, for example: ImportAll, CleanAndMap, RunValidation, BuildSubmission, SwingsReport. Export each module as a .bas file into `04_your_work` so you can show the code. Time the manual run versus the macro run and write the number down.

### Module 9 (stretch): US version
JD: *"ensuring that all US data obligations are met"*

Take your final data and produce a simplified US-style report in pounds by material. US state packaging laws each have their own categories and formats, so treat this as a format conversion exercise only, not real US rules.

## What to say in an interview

"I built a practice pipeline on real UK EPR categories: eight clients, eight different file formats, about 11,000 rows. I standardised it into one table, wrote validation rules that caught unit errors, duplicates and invalid categories, ran a swings analysis against the prior year, and automated the steps in VBA. Here is what it caught."

Only say the parts you have actually built.
