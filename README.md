# Excel Practicals: Complete Revision README

Source file: `Book1 (1).xlsx` (19 sheets). This README explains every sheet one by one: what it is for, how it was built step by step, every formula broken down, what result you get, and the traps to avoid.

> **Honest note:** the workbook has no written "aim" for any practical. I worked out the aim of each sheet from its data, formulas, pivots and charts. The formulas, ranges and results below are read directly from your file, so they are exact. The "Aim" lines are my best interpretation, so match them to your journal wording.

---

## 0. Read this first

### Sheet order and what each one covers

| # | Sheet | Topic |
|---|-------|-------|
| 1 | Sheet1 | Student marks: AVERAGE, MAX, MIN, ROUNDUP, IF, RANDBETWEEN |
| 2 | Sheet2 | Same + AND/OR, nested IF (grades), conditional formatting |
| 3 | Sheet3 | VLOOKUP, XLOOKUP, drop-down list (data validation) |
| 4 | Sheet4 | INDEX (timetable grid) |
| 5 | Sheet5 | UNIQUE + SUMIF |
| 6 | Sheet6 | UNIQUE + SUMIFS + COUNTIF |
| 7 | Sheet7 | DSUM, AVERAGEIF, AVERAGEIFS, UNIQUE, drop-downs |
| 8 | Sheet8 | Loan EMI: PMT, IPMT, PPMT |
| 9 | Sheet9 | Seat booking: COUNTA, COUNTBLANK, COUNTIF, conditional formatting |
| 10 | Sheet10 | Pivot table + pivot chart + pie chart |
| 11 | Sheet11 | Second pivot table (sales by city) |
| 12 | Sheet12 | Excel Table, OFFSET, two pivot tables |
| 13 | Sheet13 | SUM / AVERAGE on marks (+ a People Graph add-in picture) |
| 14 | Sheet14 | SEQUENCE + two Map charts |
| 15 | Sheet15 | DATEDIF, TODAY, SUMIF (HR data) |
| 16 | Sheet16 | Splitting joined names (Flash Fill style) |
| 17 | Sheet18 | Splitting "12ADI" into number + name |
| 18 | name | Mail-merge data list (name, pending amount, email) |
| 19 | Sheet17 | Second mail-merge data list |

(Sheet 17 is named `Sheet18` and sheet 19 is named `Sheet17`. The names are swapped, so go by position when someone says "sheet 17".)

### Concepts that appear everywhere

**1. Formulas always start with `=`.**

**2. Cell references**
- Relative `A1`: changes when you copy or drag the formula.
- Absolute `$A$1`: never changes (press **F4** to toggle).
- Mixed `$A1` or `A$1`: locks only the column or only the row.
- Many of your formulas "drift" when dragged, for example `SUMIF($C$2:C21,...)` becomes `C22`, `C23`... It still works because the extra cells are empty, but the correct way is `$C$2:$C$21`.

**3. `RANDBETWEEN(low, high)`** gives a random whole number. It is *volatile*: it changes every time the sheet recalculates (any edit, or **F9**). So the marks, salaries and sales in your file change every time you open it. The numbers quoted in this README are the values saved in your file. Yours will differ. To freeze them: copy → Paste Special → Values.

**4. Array formulas / dynamic arrays**
- One formula that fills a whole range. In Excel 365 it "spills" automatically.
- Old Excel needs **Ctrl+Shift+Enter**, and the formula shows in `{ }`.
- Examples in your file: `UNIQUE`, `SEQUENCE`, `ROUNDUP(F2:F21,2)`.

**5. Criteria functions** (`SUMIF`, `COUNTIF`, `AVERAGEIF`, and their `...IFS` versions) filter rows by a condition. Text criteria are not case-sensitive. **Hidden spaces break matches** (`"cash "` is not `"cash"`).

**6. Argument order trap**
- `SUMIF(range, criteria, sum_range)`: the sum range is **last**.
- `SUMIFS(sum_range, criteria_range1, criteria1, ...)`: the sum range is **first**.
- Same for `AVERAGEIF` vs `AVERAGEIFS`.

---

## 1. Sheet1: Student Marks Result

**Aim:** Calculate the average, highest, lowest, rounded average, eligibility and pass/fail for 20 students in 3 subjects.

**Columns (A1:L1):** rollno | name | phy | chem | math | avg | max | min | round off | age | criteria | result

### Steps
1. Type the headers in row 1. Enter roll numbers 1 to 20 in A and names in B.
2. **Marks:** `C2` is typed as `55` (a fixed value). All other marks (C3:C21, D2:E21) use:
   `=RANDBETWEEN(20,99)` → a random mark from 20 to 99. Fill down / across.
3. **Average (F):** `=AVERAGE(C2:E2)`. Drag down to row 21.
4. **Max (G):** `=MAX(C2:E2)`. **Min (H):** `=MIN(C2:E2)`. Drag down.
5. **Round off (I):** one array formula over `I2:I21`:
   `=ROUNDUP(F2:F21,2)`
   - `ROUNDUP(number, digits)` always rounds *away from zero*. Example: 46.3333 becomes **46.34**. (`ROUND` would give 46.33.)
   - Non-array alternative: `=ROUNDUP(F2,2)` and drag down.
6. **Age (J):** `=RANDBETWEEN(14,28)`.
7. **Criteria (K):** one array formula over `K2:K21`:
   `=IF(J2:J21>18,"qualify","not applicable")`
   Age above 18 gives "qualify", otherwise "not applicable".
8. **Result (L):** `=IF(I2>=55,"pass","fail")`
   Pass mark is **55**, tested on the *rounded average*.

### Formula meanings
| Formula | Meaning |
|---|---|
| `IF(test, value_if_true, value_if_false)` | Does one thing if the test is true, another if not |
| `AVERAGE(range)` | Mean of the numbers |
| `MAX` / `MIN` | Largest / smallest |

### Sample result from your file (row 2)
phy 55, chem 23, math 61 → avg 46.33 → rounded 46.34 → age 22 → "qualify" → **fail** (46.34 is below 55).

**Exam tips:** the results change on every recalculation because of RANDBETWEEN, so mention that. Age 18 exactly gives "not applicable" because the test is `>18`, not `>=18`.

---

## 2. Sheet2: Marks + Grades + Colour Rules

**Aim:** Same marks table (10 students, dish names as names), plus subject-wise pass checks, grades and colour highlighting.

Columns A to L are built exactly like Sheet1 (avg, max, min, round off, age, criteria, result). `C2` is typed as 98. Extra columns:

### M: "result1" (must pass ALL subjects)
`=IF(AND(C2>33,D2>33,E2>33),"pass","fail")`
- `AND(...)` is TRUE only if **every** condition is true.
- A student passes only if each subject is above 33.

### N: "grade1" (fail if ANY subject is low)
`=IF(OR(C2<33,D2<33,E2<33),"fail","pass")`
- `OR(...)` is TRUE if **at least one** condition is true.
- Note the logic is flipped: "if any subject is below 33 → fail".
- **M and N are the same idea written two ways.** They differ only at exactly 33: M fails a 33 (`>33` is false), N passes it (`<33` is false).

### O: "Grade" (nested IF)
`=IF(F2>=90,"A",IF(F2>=50,"B",IF(F2>=10,"C")))`
- Average 90 or above is **A**; 50 to 89 is **B**; 10 to 49 is **C**.
- Trap: the last `IF` has no "else" part, so an average below 10 returns `FALSE`. A safer version: `=IF(F2>=90,"A",IF(F2>=50,"B","C"))`.
- Nested IF is read left to right and stops at the first true test, so the order of conditions matters.

### Conditional formatting on C2:E11 (the marks)
Home → Conditional Formatting → Highlight Cells Rules → Less Than / Greater Than. Rules in priority order:

| Priority | Rule | Look |
|---|---|---|
| 1 | less than 36 | light red fill |
| 2 | greater than 85 | green |
| 3 | greater than 55 | yellow |
| 4 | greater than 35 | yellow |
| 5 | less than 33 | red text + red fill |
| 6 | greater than 80 | yellow |

Several rules can apply to one cell. When their formats clash, the higher-priority (lower number) rule wins. Manage them from Conditional Formatting → Manage Rules.

---

## 3. Sheet3: VLOOKUP, XLOOKUP and a Drop-down

**Aim:** Look up an employee's details from a table of 15 employees (sr no, name, designation, salary).

Salary column: `=RANDBETWEEN(10000,90000)`.

Lookup panel (F2:G4):

| Cell | Formula | Returns |
|---|---|---|
| G2 | `=VLOOKUP(B2,B1:D16,1,0)` | the name itself (column 1 of the table) |
| G3 | `=XLOOKUP(B2,B1:B16,C1:C16)` | the designation |
| G4 | `=VLOOKUP(B2,B1:D16,3,0)` | the salary (column 3 of B:D) |

### VLOOKUP
`VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
- `B2` is the value to find.
- `B1:D16` is the table. **The lookup value must be in the first column of the table** (column B).
- Column index: B=1, C=2, D=3. So **3 gives salary**.
- `0` (or FALSE) means **exact match**. Always use 0 for names.
- VLOOKUP can only look to the **right** of the lookup column.

### XLOOKUP (Excel 365)
`XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`
- Looks in one column (`B1:B16`) and returns from another (`C1:C16`).
- Exact match by default, can look left, and no column numbers to count.

### Drop-down list (cell H7)
Data → Data Validation → Allow: **List** → Source: `=$B$2:$B$16`.
Result: H7 shows a drop-down of the 15 names (currently "vidya").

**Gotcha in your file:** the formulas in G2:G4 look up `B2`, not `H7`, so the drop-down is not connected. To make it work, replace `B2` with `H7` in all three formulas. Then picking a name in H7 changes the results.

---

## 4. Sheet4: INDEX on a Timetable Grid

**Aim:** Show a timetable-style grid (time slots × subjects) and pick a value by row and column position.

- Column A holds times shown as `3:36, 4:36, 5:36 ... 12:36` (one hour apart). The cells hold fractions of a day, formatted as time `h:mm`.
- Columns B to E hold subjects (math, chem, bio, eco, ...), one row per time slot.
- Cell `I9`: `=INDEX(A1:E11,2,2)` returns **math**.

### INDEX
`INDEX(array, row_num, [column_num])`
- Counts **inside the selected range**: row 2, column 2 of `A1:E11` is cell **B2**.
- Do not count from the sheet's A1. Count from the first cell of the range you selected.

**Exam tip:** INDEX is often paired with MATCH for two-way lookups: `=INDEX(range, MATCH(row_item, row_headers, 0), MATCH(col_item, col_headers, 0))`. That pairing is not in your file, but it is a common question.

---

## 5. Sheet5: UNIQUE + SUMIF (salary per designation)

**Aim:** Get the list of distinct designations and total salary for each.

Data (A1:D21): sr no, name, designation, salary (typed values) for 20 employees. Note `CFO` appears twice (devarsh and shlok).

| Cell | Formula | What it does |
|---|---|---|
| E2 | `=UNIQUE(D2:D21)` | spills all distinct salaries down column E |
| F2 | `=UNIQUE(C2:C21)` | spills the distinct designations (19 of them, because CFO repeats) |
| G2 | `=SUMIF($C$2:C21,F2,$D$2:D21)` | total salary for the designation in F2 |

`G2` is filled down along the spill.

### SUMIF
`SUMIF(range, criteria, [sum_range])`
- `range` is where to check (designations, C).
- `criteria` is what to match (`F2`).
- `sum_range` is what to add (salary, D).

**Result highlight:** CFO = 74,169 + 86,861 = **161,030**. The other designations show their single salary. `G21` shows 0 because F21 is empty.

**Gotcha:** the end of the range drifts (`C21`, `C22`, `C23`...) as the formula is filled down. It works only because those extra cells are empty. Lock it as `$C$2:$C$21` and `$D$2:$D$21`.

---

## 6. Sheet6: UNIQUE + SUMIFS + COUNTIF

**Aim:** Total salary by occupation and marital status, plus counts.

Data A1:D20: name, occupation, marital status, salary (`=RANDBETWEEN(20000,150000)`).

### Salary matrix (G2:I9)
- `G3`: `=UNIQUE(B2:B20)` spills the 7 distinct occupations (G3:G9).
- Headers: `H2` = married, `I2` = unmarried.
- `H3` (filled across and down):
  `=SUMIFS($D$2:D20,$B$2:B20,$G3,$C$2:C20,$H$2)`
  - Adds salary where **occupation = G3 AND marital status = H2**.
  - `$G3` locks the column (occupation list), `$H$2` locks the header cell.
- `I3` is the same with `$I$2`.

### Counts
| Cell | Formula | Result |
|---|---|---|
| G11 (married) | `=COUNTIF(C2:C20,C4)` | counts cells equal to "married" (C4 holds "married") |
| G12 (unmarried) | `=COUNTIF(C2:C20,C2)` | counts cells equal to "unmarried" |
| G13:G19 | `=COUNTIF(B2:B20,$F$13:F19)` (array) | head-count per occupation (officer, accountant, CEO, scientist, engineer, clerk, teacher) |

### Data error to know about
Cell `C3` (Aarna) says **"unmerried"** (misspelt). Excel treats it as a different word, so:
- the *married* count is 8 and the *unmarried* count is 10, which adds up to 18, not 19;
- Aarna's salary (an accountant) is missing from the "unmarried" total for accountants, so that total is short by her salary.

Fix the spelling and everything updates.

### SUMIFS vs SUMIF
`SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)`. All the conditions must be true (AND logic).

---

## 7. Sheet7: Sales Data: DSUM, AVERAGEIF, AVERAGEIFS

**Aim:** Analyse Apple-product sales (19 rows) by device and region using database and conditional-average functions.

Data A1:E20: device | region | month | party name | sale.
Helper lists:
- `F2`: `=UNIQUE(B2:B20)` spills the 4 regions (east, west, south, north).
- `G2`: `=UNIQUE(A2:A20)` spills the 15 distinct devices.

### Drop-downs for the DSUM criteria
- `I4` (region): Data Validation → List → `=$F$2:$F$5`
- `J4` (device): Data Validation → List → `=$G$2:$G$20`

### DSUM: `K4`
`=DSUM(A1:G20,E1,I3:J4)`

`DSUM(database, field, criteria)`
- `database` = `A1:G20`: the whole table **including the header row**.
- `field` = `E1`: which column to add (the header "sale").
- `criteria` = `I3:J4`: a mini table. Row 3 holds the **headers** (`region`, `device `) and row 4 holds the **conditions** (`south`, `macbook`).
- Conditions on the same row are AND. So this adds sales where **region = south AND device = macbook** → **116,055**.
- **Trap:** the header in the criteria must match the database header *exactly*. Your headers are `device ` with a trailing space in both places. That is why it works.

### AVERAGEIF: `J6:J24`
`=AVERAGEIF(A1:A20,I6,E1:E20)`
`AVERAGEIF(range, criteria, average_range)`: the average sale of the device named in `I6` (a typed list of devices). Example: macbook → average of 116,055 and 107,063 = **111,559**.

### AVERAGEIFS: `K6:K24`
`=AVERAGEIFS(E2:E20,B2:B20,B2,A2:A20,A2)`
`AVERAGEIFS(average_range, criteria_range1, criteria1, criteria_range2, criteria2, ...)`: the average sale where **region = B2 AND device = A2**. Each row looks at one data row's region and device pair. Because there is one record per combination here, the result equals that sale.

**Notice:** AVERAGEIF puts the average range **last**; AVERAGEIFS puts it **first**. The ranges drift as the formulas are filled down, which works only because the ranges are long enough.

---

## 8. Sheet8: Loan EMI (PMT, IPMT, PPMT)

**Aim:** Calculate the monthly EMI for a loan and split each EMI into interest and principal.

Inputs:
| Cell | Value |
|---|---|
| C1 | interest rate = 0.1 (10% per year) |
| C2 | amount = 100,000 |
| C3 | duration = 10 months |

### EMI: `C4`
`=PMT(C1/12,C3,-C2,,1)` → **10,377.56**

`PMT(rate, nper, pv, [fv], [type])`
- `rate` = `C1/12`: **annual rate divided by 12**, because payments are monthly.
- `nper` = `C3`: number of payments.
- `pv` = `-C2`: the loan is entered as a **negative** so the EMI comes out positive.
- `fv` is skipped (empty comma).
- `type` = `1`: payment is made at the **start** of each month. (`0` or omitted means end of month.)

### Total paid: `C6`
`=(C4*C3)` → 103,775.58. Total interest = 103,775.58 − 100,000 = **3,775.58**.

### Interest and principal split (columns F, I, J)
- `F2:F10` hold the period numbers 2 to 10 (and `H1` has a stray `1`).
- `I1:I10`: array `=IPMT(C1/12,F1:F10,C3,-C2)` is the **interest part** of each EMI.
- `J1:J10`: array `=PPMT(C1/12,F1:F10,C3,-C2)` is the **principal part** of each EMI.

`IPMT(rate, per, nper, pv, [fv], [type])` and `PPMT(...)` use the same arguments with `per` = which period.

### Two real problems in this sheet
1. **`I1` and `J1` show `#NUM!`.** The period range starts at `F1`, which is empty, and period 0 is invalid. Fix: put `1` in `F1` (and delete the stray `H1`), or start the range at `F2`.
2. **Mismatch with type:** `PMT` uses `type=1` (start of month) but `IPMT`/`PPMT` don't pass a type, so they assume end of month. As a result interest + principal per row (about 10,464.04) does not equal the EMI in C4 (10,377.56). Use the same `type` everywhere. Add `,,1` to IPMT/PPMT, or drop `,,1` from PMT.

### Unfinished block
`B14:C16` has `emi = 15,000`, `60`, `0.12` with no formula. This looks like the setup for the reverse question: how big a loan can a 15,000 EMI repay over 60 months at 12%? Formula (my addition, not in your file): `=PV(0.12/12,60,-15000)` ≈ **6,74,325**.

---

## 9. Sheet9: Seat Booking Grid (COUNT functions)

**Aim:** A seating chart (7 columns × 16 rows) where `m` = male booking, `f` = female booking, blank = free. Count bookings.

Counts (K5:L8):

| Label | Formula | Result |
|---|---|---|
| booked | `=COUNTA(A1:G16)` | **47** |
| remaining | `=COUNTBLANK(A1:G16)` | **65** |
| male | `=COUNTIF(A1:G16,"m")` | **24** |
| female | `=COUNTIF(A1:G16,"f")` | **22** |

- `COUNTA` counts **non-empty** cells. `COUNTBLANK` counts **empty** cells. Together: 47 + 65 = 112 = 7 × 16 seats. ✔
- `COUNTIF(range, criteria)` counts cells equal to the criteria. Text goes in quotes.
- `COUNT` (not used here) counts only numbers.
- **Data error:** cell `G12` contains `fs` (a typo). It is counted as booked (47) but is neither "m" nor "f", so 24 + 22 = 46, one short of 47.

### Conditional formatting on A1:G16
Four "Cell Value" rules compare text alphabetically: greater than "m", less than "f", greater than "m", greater than "f". Text comparison is alphabetical, so "m" is greater than "f". That is how `m` and `f` seats get different colours. Open Manage Rules to see the exact colours.

---

## 10. Sheet10: Orders Data + Pivot Table + Charts

**Aim:** Summarise sales by country and region with a pivot table, and show it as charts.

Data A1:I16 (15 orders): order no, order date, customer name, country, state, city, region, sales (`=RANDBETWEEN(1000,4000)`), profit (`=RANDBETWEEN(1500,3000)`).

**Note:** order dates (like `01.01.23`) are stored as **text**, not real dates. Date functions and date grouping will not work on them until converted.

### Pivot table (B19:H26): "PivotTable1"
1. Click inside the data → **Insert → PivotTable** → source `Sheet10!A1:I16` → place it in the same sheet at B19.
2. Drag **country** to **Rows**.
3. Drag **region** to **Columns**.
4. Drag **sales** to **Values** (set to **Sum**) → the pivot shows "Sum of sales".

Result saved in your file: countries down the side, regions (abz, bdh, fcki, hsb, ncs) across the top, Grand Total **41,588** (for example india = 11,010).

### Charts on this sheet
- **Pivot chart (clustered column):** select the pivot → Insert → Column chart → Clustered Column. It uses the pivot (rows `B21:B26`, one series per region).
- **Pie chart:** categories `D2:D16` (country), values `H2:H16` (sales). Select those columns → Insert → Pie. Each order is one slice.

### Pivot rules to remember
- A pivot **does not update itself**. After the data changes (and RANDBETWEEN changes it constantly), right-click → **Refresh** (or Alt+F5).
- Pivot totals can therefore disagree with the live data.

---

## 11. Sheet11: Pivot Table: Sales by City

**Aim:** A second pivot built from the same orders data.

- Source: `Sheet10!A1:I16`.
- Location: `A3:B9`.
- Rows: **city**, with **order date** as a second row field (only the city level is shown).
- Values: **Sum of sales**.

Saved result: ahd 10,075 | denpasar 6,850 | illinois 4,984 | london 7,982 | melbourne 8,417 | Grand Total **38,308**.

**Why 38,308 and not 41,588 (Sheet10)?** Each pivot made separately from the same range gets its own cache, and the sales are random numbers, so the two pivots were last refreshed at different moments. (This is my reading. Refresh both and compare to verify.)

---

## 12. Sheet12: Excel Table, OFFSET and Pivot Tables

**Aim:** Convert data to a Table, copy it with a dynamic OFFSET formula, and summarise revenue by customer type.

Data A1:C14: zone | custome type | revenue (for example north / cash / 2448).

### Excel Table
Select `A1:C14` → **Ctrl+T** (Insert → Table). Your file names it **Table1**, style *Medium 2*, with filter buttons. Tables grow automatically when you add rows.

### OFFSET array (E2:G15)
`=OFFSET($A$2,0,0,COUNTA($A:$A),COUNTA($1:$1))`

`OFFSET(reference, rows, cols, [height], [width])`
- Start at `A2`, move 0 rows and 0 columns.
- **height** = `COUNTA($A:$A)` = number of filled cells in column A (14, includes the header).
- **width*
