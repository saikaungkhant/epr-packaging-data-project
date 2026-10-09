# Every problem at a glance

A quick list of everything that went wrong and how I found and fixed it. Each one is a possible interview story.

| # | Module | What went wrong | How I found it | Fix | Lesson |
|---|---|---|---|---|---|
| 1 | 2 | Row count 10,266, then 20,135, instead of 10,248 | Per-client COUNTIF table vs expected counts | Deleted 5 empty BBO lines and about 9,900 unlabelled rows | Reconcile per group, not just the total |
| 2 | 3 | Find & Replace removed every comma from every formula | All the lookups broke | Restored from version history; moved comma handling into the formula | Never bulk-edit; put cleaning logic in formulas |
| 3 | 3 | Every row showed weight 23 | Formula bar showed J2 and K2 on every row | Cleared and filled down properly from row 2 | Fill down, don't paste formula text |
| 4 | 4 | False duplicates (rows with different units) | Checked flagged pairs by eye: 3,252 vs 2,320 units | Added units to the key | A key must include every field that makes a row unique |
| 5 | 4 | The fixed key still didn't work | The new formula was only on visible rows | Cleared filters, filled down again: 12 true duplicates | Never fill down with a filter on |
| 6 | 4 | Average flagged the wrong rows in the unit-weight check | Tested average vs median before using either | Used the median per group | One extreme value drags an average |
| 7 | 5 | 26 SKUs overwritten with AVD-0002 | COUNTIF gave 72 instead of 46 | Looked up the 26 correct SKUs, retyped one by one, proved with counts that should be 0 | With a filter on, type one cell at a time |
| 8 | 5 | 4 real BBO rows deleted one row off | BBO count 2,379 vs 2,383; expected-rows list plus COUNTIFS anti-join | Re-added the 4 exact rows | Delete through a filter, never by eye |
| 9 | 5 | Pasting the 4 rows overwrote 3 CBH rows | CBH count dropped to 1,917 | Same anti-join, re-added, pasted below Ctrl+End | Paste always overwrites |
| 10 | 5 | 43 missing weights but only 37 in the replies | Broke 43 down by client | 25 typed overrides plus 18 CCR rows rebuilt by the re-join | Match on key fields, never on position |
| 11 | 5 | 752 UNMAPPED after adding SPF's late file | 752 matched the new rows exactly; LEN 7, COUNTIF 0 | Added the "2025-H2" mapping, switched to whole-column ranges | Test one layer at a time |
| 12 | 6 | Every submission weight was 0 | ISNUMBER(AE2) was FALSE; final_kg empty | Re-entered the final_kg formula | Check the column, then the key, then the value |
| 13 | 6 | 440 rows vs 446, same total kg | XLOOKUP vs answer key: 8 MDC keys missing, all P2 to P4 | MDC class formula pointed at row 5; re-entered it | Totals match but counts don't means a labelling problem |
| 14 | 6 | Missed MDC's typed total (41.5 t wrong) | The answer key's issues log | Avoided only by luck (never used TOTAL) | Recalculate typed totals and ask about gaps |
| 15 | 7 | Called real swings "errors" and doubted a 31% rise | Worked through the maths and the client replies | Kept the data, wrote the reason | A flag is a question, not a verdict |
| 16 | 7 | AVD glass -19% not flagged | Read the rows just under the threshold | Noted it as linked to the aluminium swing | A threshold is a tool, not a rule |
| 17 | 8 | Recorder captured my Save As | Read the recorded code line by line | Rewrote the macro by hand | Always read and clean recorded code |
| 18 | 8 | Button: "Cannot run Button1_Click" | Error message | Re-assigned by clicking RunChecks in the list | Pick the macro from the list |
| 19 | 8 | F5 opened Go To; message seemed to vanish | Tried it | F5 only in the editor, Alt+F8 from Excel; message was behind the editor | Know where the code runs |
