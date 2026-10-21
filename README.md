# Excel Practical Exam: Revision Guide

> Built from your own workbook (`Book1.xlsx`, 25 sheets). Every section tells you **what you did, how the formula works, where you slipped, and a quick self-test.** I've read your formulas, pivots, charts and dropdowns, and flagged real mistakes I found in the file so you don't repeat them in the exam.

**Legend:** [365] = Excel 365/2021 only (UNIQUE, XLOOKUP, SEQUENCE, spilling). Check the exam PC has it. [MISTAKE] = a mistake found in your sheet. [EXAM] = likely exam question.

---

## 0. Instructor's Golden Rules (read these twice)

1. **Press `F4` to lock references.** `C2:C21` → `$C$2:$C$21`. This is the most common mistake in your file (see Sheets 5, 6, 7, 15).
2. **Know your argument order.** `SUMIF(range, criteria, sum_range)` but `SUMIFS(sum_range, range1, crit1, ...)`. Sum-range moves to the front in the "S" versions.
3. **`0` or `FALSE` as the last VLOOKUP argument means exact match.** Always use it.
4. **Always give nested IFs a final "else".** Yours returned `FALSE` (Sheet1).
5. **`RANDBETWEEN` and `TODAY` are volatile.** They change on every recalculation. Freeze values with Copy → Paste Special → Values (`Ctrl+Alt+V`, then `V`).
6. **Pivot tables do not auto-update.** Right-click → Refresh (`Alt+F5`).
7. **Dates typed as `10.10.2026` stay as text.** They won't group or calculate. Use real dates (`10/10/2026`).
8. **Spelling matters.** `QAULIFY`, `UNQULIFY`, `nevember` and `cicago` all appear in your sheets. Examiners check output text exactly.

---

## 1. Sheet Map

| Sheet | Topic | Key skills |
|---|---|---|
| Sheet1 | Student marks | AVERAGE, MAX, MIN, IF, AND, nested IF, conditional formatting |
| Sheet2 | Ages | UNIQUE [365], FREQUENCY |
| Sheet3 | Employee lookup | VLOOKUP, XLOOKUP [365], drop-down list |
| Sheet4 | Timetable | INDEX, drop-down list |
| Sheet5 | Salary by designation | SUMIF + UNIQUE |
| Sheet6 | Salary × marital status | SUMIFS (2 conditions), COUNTIF |
| Sheet7 | Device sales | DSUM, AVERAGEIF, AVERAGEIFS, dependent drop-downs |
| Sheet8 | Loan | PMT, IPMT, PPMT |
| Sheet9 | Seat booking | COUNTA, COUNTBLANK, COUNTIF, conditional formatting, colour scales |
| Sheet10 | Orders | Pivot table, pivot chart, pie chart |
| Sheet11 | Revenue | Excel Table (`Ctrl+T`), 2 pivots, OFFSET |
| Sheet12 | City sales | Pivot with 2 row fields |
| Sheet13 | Marks totals | SUM, AVERAGE, pictures, Insert-tab tour |
| Sheet14 | Number grid + countries | SEQUENCE [365], filled map chart |
| Sheet15 | Staff | DATEDIF, TODAY, wildcard SUMIF |
| Sheet16 | Names | Custom function `AddNameSpaces` (VBA) |
| Sheet17 | Text split | Custom functions `GETNAME`, `GetNumbers` (VBA) |
| Sheet18 + 4 city sheets | Customers | AutoFilter, splitting data onto separate sheets |
| Sheet23 | State/category sales | Raw data for pivot/chart practice |
| Sheet20 | Pending payments | E-mail hyperlinks |
| Sheet19 | Empty | (scratch sheet) |

---

## 2. Sheet-by-Sheet Walkthrough

### Sheet1: Student Marks (IF family + Conditional Formatting)

**Layout:** Roll No, Name, PHY, CHE, MATHS, then calculated columns.

| Col | Formula (row 2) | Meaning |
|---|---|---|
| C, E | `=RANDBETWEEN(20,99)` | Random whole number from 20 to 99 (dummy marks) |
| F | `=AVERAGE(C2:E2)` | Mean, formatted to 2 decimals (`0.00`) |
| G | `=MAX(C2:E2)` | Highest mark |
| H | `=MIN(C2:E2)` | Lowest mark |
| I | `=RANDBETWEEN(15,20)` | Random age |
| J | `=IF(I2>=18,"QUALIFY","UNQUALIFY")` | Simple IF: age 18+ qualifies |
| K | `=IF(F2>=55,"PASS","FAIL")` | Pass if average is at least 55 |
| L | `=IF(AND(C2>35,D2>35,E2>35),"pass","fail")` | Pass only if **all** subjects are above 35 |
| M | `=IF(F2>80,"A",IF(F2>70,"B",IF(F2>60,"C","FAIL")))` | Nested IF grade (corrected version) |

**Conditional formatting you applied** (Home → Conditional Formatting → Highlight Cells Rules):
- Cells `< 35` → light red fill (fail marks)
- Cells `> 90` → yellow fill (top marks)

**How IF works:** `IF(test, value_if_true, value_if_false)`. In a nested IF, the false branch contains another IF. Excel checks top to bottom and stops at the first TRUE, so **order matters**: test the highest threshold first.

**AND vs OR:** `AND` needs every condition TRUE. `OR` needs at least one.

[MISTAKE] **Mistakes in your sheet:**
- **Column M** has no final else: `IF(F2>60,"c")` returns `FALSE` when the average is 60 or below. Add `,"FAIL"` at the end.
- **Cell J5** refers to `I2` instead of `I5` (a wrong row). Row 5 shows `QAULIFY` although its age is 16.
- **Column D** contains the text `Invalid Email` (leftover, not marks).
  - `AVERAGE`, `MAX` and `MIN` **ignore text**, so your averages only used 2 subjects.
  - In `L`, `"Invalid Email">35` evaluates to **TRUE**, because Excel treats text as larger than any number. That's why `L` showed `pass`.
- Typos: `QAULIFY`, `UNQULIFY`, `UNQAULIFY`, and `c` in lowercase.

> [EXAM] **Self-test:** *Write a formula giving "Distinction" if average ≥ 75, "First" if ≥ 60, "Pass" if ≥ 40, otherwise "Fail".*
> **A:** `=IF(F2>=75,"Distinction",IF(F2>=60,"First",IF(F2>=40,"Pass","Fail")))`

---

### Sheet2: UNIQUE and FREQUENCY [365]

**Data:** 15 people with ages in `C2:C16`.

- `D2: =UNIQUE(C2:C16)` lists each distinct age once. It **spills** down to `D2:D10` (9 unique ages: 25, 28, 29, 35, 34, 26, 31, 30, 33).
- `E2: =FREQUENCY(D2:D10,C2:C16)` was meant to count how many people have each age.

**FREQUENCY syntax:** `FREQUENCY(data_array, bins_array)`
- **data_array** = all the raw values (the ages)
- **bins_array** = the cut-off points (the unique ages)
- It returns **one more result than there are bins** (the last is "values above the highest bin"), so select/spill 10 cells for 9 bins.

[MISTAKE] **Your arguments are swapped.** Correct version:

```excel
=FREQUENCY(C2:C16, D2:D10)
```

Correct counts: 25→2, 28→2, 29→3, 35→1, 34→3, 26→1, 31→1, 30→1, 33→1 (total 15).

**Easier alternative:** `=COUNTIF($C$2:$C$16, D2)` and fill down.

> [EXAM] **Self-test:** *How do you get a list of distinct ages?* → `=UNIQUE(C2:C16)`. *How do you count each?* → `FREQUENCY` or `COUNTIF`.

---

### Sheet3: Lookups + Drop-down

**Data:** Sr no, name, designation, salary (`A1:D16`). Cell `G8` has a **drop-down** of names.

**Creating the drop-down:** select `G8` → Data → Data Validation → Allow: **List** → Source: `=$B$2:$B$16`.

| Formula | Meaning |
|---|---|
| `=VLOOKUP(B2,B1:D16,1,0)` | Looks up B2 in the first column of the table and returns column 1 (the name itself) |
| `=VLOOKUP(B2,B1:D16,3,0)` | Same lookup, returns column **3** of `B:D` (the salary) |
| `=XLOOKUP(B2,B1:B16,C1:C16)` [365] | Finds B2 in `B1:B16`, returns the matching value from `C1:C16` (designation) |

**VLOOKUP anatomy:** `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])`
- The lookup column must be the **left-most** column of the table.
- `col_index_num` counts from the left of `table_array`. In `B:D`: name = 1, designation = 2, salary = 3.
- Last argument `0` = exact match.

**XLOOKUP anatomy:** `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`. It can look left, has a built-in "not found" message, and defaults to exact match.

[MISTAKE] **What to fix:** the lookup value was fixed at `B2` and the table wasn't locked. The idea of the drop-down is to look up **the chosen name**:

```excel
=VLOOKUP($G$8, $B$1:$D$16, 3, 0)        ← salary of whoever is selected
=XLOOKUP($G$8, $B$2:$B$16, $C$2:$C$16)  ← designation of whoever is selected
```

> [EXAM] **Self-test:** *Salary of "tanvi" with VLOOKUP?* → `=VLOOKUP("tanvi",B2:D16,3,0)`. *Why `#N/A`?* The value isn't found (typo, trailing space, or wrong lookup column).

---

### Sheet4: INDEX and Drop-down

**Data:** A timetable. Times in `A2:A11` (11.15, 12.15, …) and columns `a`, `b`, `c`, `d` hold subjects. `H6` has a drop-down of times.

- `I6: =INDEX(A1:E11,2,3)` gives **`chem`**: row 2, column 3 of that block (cell `C2`).

**INDEX syntax:** `INDEX(array, row_num, column_num)` returns the value at that position inside the range.

**Power move: INDEX + MATCH** (what you'd use with the drop-down):

```excel
=INDEX($B$2:$E$11, MATCH($H$6,$A$2:$A$11,0), MATCH("c",$B$1:$E$1,0))
```
- `MATCH(value, range, 0)` returns the **position** of the value.
- Together: "go to the row where time = H6, and the column headed c."

> [EXAM] **Self-test:** *What does `=INDEX(A1:E11,5,2)` return?* → Row 5 of the block is the 14.15 row, column 2 = `org chem` (cell `B5`).

---

### Sheet5: SUMIF with UNIQUE [365]

**Data:** 20 employees with designation (`C`) and salary (`D`).

- `E2: =UNIQUE(C2:C21)` spills the designation list (14 designations).
- `F2: =SUMIF(C2:C21,E2,D2:D21)` is the total salary per designation.

**SUMIF anatomy:** `SUMIF(range_to_check, criteria, range_to_add)`.

[MISTAKE] **Drifting ranges.** You filled down without locking, so the ranges slid:
`F2` uses `C2:C21`, `F6` uses `C6:C25`, `F10` uses `C10:C29`, `F14` uses `C14:C33`.
The answers happened to be right only because the matching rows were still inside the window. Always write:

```excel
=SUMIF($C$2:$C$21, E2, $D$2:$D$21)
```

> [EXAM] **Self-test:** *Total salary of all "manager" staff?* → `=SUMIF(C2:C21,"manager",D2:D21)`

---

### Sheet6: SUMIFS (two conditions) and COUNTIF

**Data:** designation (`C`), salary (`D`), marital status (`E`, with `m` = married, `u` = unmarried). A summary grid sits at `H2:J8`.

```excel
=SUMIFS($D$2:$D$21, $C$2:$C$21, $H3, $E$2:$E$21, I$2)
```
- Sum salary where designation = the row label **and** status = the column header.
- `$H3` locks the **column** only (so it can fill right) and `I$2` locks the **row** only (so it can fill down). This is a **mixed reference**. One formula fills the whole grid.

**Your version:** `$D$2:D21` locks only the start. When copied down it grows to `D22`, `D23`... It still worked because the extra rows are blank, but write `$D$2:$D$21`.

**COUNTIF block:**
- `=COUNTIF(E2:E21,"m")` → 9 married; `"u"` → 11 unmarried (9 + 11 = 20 OK)
- `=COUNTIF(C2:C21,H13)` → number of staff per designation

[MISTAKE] `H13: =UNIQUE(C12:C31)` includes blank cells below the data, so a stray `0` appears in `H19`. Use `UNIQUE(C2:C21)`.

**Merged cell:** `H10:I10` ("count") was merged with Home → Merge & Center.

> [EXAM] **Self-test:** *Salary total of unmarried directors?* → `=SUMIFS(D2:D21,C2:C21,"director",E2:E21,"u")`

---

### Sheet7: DSUM, AVERAGEIF, AVERAGEIFS, drop-downs

**Data (`A1:E20`):** device, region, month, party name, sale. Helper lists: `F` = unique regions, `G` = unique devices (both from `UNIQUE`).

**Drop-downs:** `I2` (region, source `$F$2:$F$5`) and `J2` (device, source `$G$2:$G$10`).

**DSUM (database function)**: `=DSUM(A1:G20, E1, I1:J2)`
- `database` = whole table **including headers**
- `field` = which column to add (`E1` = "sale")
- `criteria` = a small block: **header row + condition row**. Here `region | device` over `east | I tag`.
- Conditions on the **same row = AND**; conditions on **different rows = OR**.
- Criteria headers must match the database headers **exactly**.

**AVERAGEIF** `(range, criteria, average_range)` → average sale per device:
`=AVERAGEIF(A1:A20, I6, E1:E20)`

**AVERAGEIFS** `(average_range, crit_range1, crit1, crit_range2, crit2…)`:
`=AVERAGEIFS(E1:E20, B1:B20, B2, A1:A20, A2)` → average where region = B2 **and** device = A2.

[MISTAKE] **Problems:**
- AVERAGEIF ranges drift down (`A1:A20`, `A2:A21`, `A3:A22`…). Lock them: `$A$2:$A$20`, `$E$2:$E$20`.
- Column K uses relative `B2`/`A2`, so each row averages its own region+device pair, which isn't a meaningful summary. For the **drop-down selection**, point at `$I$2` and `$J$2`:
  `=AVERAGEIFS($E$2:$E$20,$B$2:$B$20,$I$2,$A$2:$A$20,$J$2)`
- Data typos: `ipade pro` vs `ipad pro` become two "unique" devices. **Clean your data first** (`TRIM`, spell-check).

> [EXAM] **Self-test:** *Total sale for "south" region and "iphone air" with DSUM?* → Criteria block `region | device` / `south | iphone air`, then `=DSUM(A1:E20,"sale",I1:J2)`.

---

### Sheet8: Loan EMI (PMT, IPMT, PPMT)

**Inputs:** rate `C1` = 10% (yearly), amount `C2` = ₹100,000, duration `C3` = 10 months.

| Cell | Formula | Meaning |
|---|---|---|
| `C4` | `=PMT(C1/12,C3,-C2,,1)` | Monthly EMI (₹10,377.56) |
| `C6` | `=C4*C3` | Total paid over the loan (₹103,775.58) |
| `G1:G10` | `=IPMT(C1/12,F1:F10,C3,-C2)` | **Interest** part of each month's EMI |
| `H1:H10` | `=PPMT(C1/12,F1:F10,C3,-C2)` | **Principal** part of each month's EMI |

**Syntax:**
- `PMT(rate, nper, pv, [fv], [type])`
- `IPMT(rate, per, nper, pv, [fv], [type])`
- `PPMT(rate, per, nper, pv, [fv], [type])`

**Key ideas:**
- Divide the yearly rate by 12 for monthly payments. `nper` is in months.
- Put `-` in front of PV so the EMI comes out **positive**.
- `type`: `0` = payment at end of month (default), `1` = start of month.
- **EMI = Interest + Principal** every month. Interest falls each month while principal rises (your column G decreases from ~₹833 to ~₹86, and H increases).

[MISTAKE] `PMT` uses `type = 1` but `IPMT`/`PPMT` use the default `type = 0`, so `G + H` won't equal `C4`. Use the **same type** everywhere (normally leave it out, giving 0).

> [EXAM] **Self-test:** *EMI on ₹5,00,000 at 12% p.a. for 24 months?* → `=PMT(12%/12,24,-500000)`

---

### Sheet9: Seat Booking (COUNT family + Conditional Formatting)

A 12-row × 7-column seat grid where `m` = male and `f` = female; blanks are free seats (`A1:G12`).

| Label | Formula | Result |
|---|---|---|
| booked | `=COUNTA(A1:G12)` | 33 (non-empty cells) |
| remaining | `=COUNTBLANK(A1:G12)` | 51 (empty cells) |
| men | `=COUNTIF(A1:G12,"m")` | 19 |
| female | `=COUNTIF(A1:G12,"f")` | 14 |

Check: 33 + 51 = 84 = 12 × 7 OK and 19 + 14 = 33 OK

**Differences:** `COUNT` counts **numbers only**, `COUNTA` counts **anything non-empty**, `COUNTBLANK` counts **empty** cells, and `COUNTIF` counts cells **meeting a condition**.

**Conditional formatting:** `f` → yellow, `m` → red (Highlight Cells Rules → *Equal To…*). You also tried **colour scales** on the number block `P18:S21` (Home → Conditional Formatting → Color Scales).

[MISTAKE] `J6` uses `A1:G13` (one row too many). Keep all four ranges identical. Also stray rules (`greater than "m"`) exist and can be deleted via Manage Rules.

> [EXAM] **Self-test:** *What percentage of seats are booked?* → `=COUNTA(A1:G12)/(ROWS(A1:G12)*COLUMNS(A1:G12))`

---

### Sheet10: Pivot Table + Charts

**Data (`A1:I16`):** order no, date, customer, country, state, city, region, sales, profit.

**Pivot table (`A18:F24`)**: Rows = **country**, Columns = **region**, Values = **Sum of sales**.

**How to make one:** click inside the data → Insert → PivotTable → OK → drag fields into **Rows / Columns / Values / Filters**.

**Charts:** a 3-D column/bar chart and a pie chart on this sheet. One is a **PivotChart** linked to the pivot table, and the other plots region vs sales from `G2:H16`. Insert → Charts, or Insert → PivotChart.

[MISTAKE] **Traps:**
- Dates such as `10.10.2026` are **text** (mixed `dd/mm/yyyy` and `mm-dd-yy` formats applied), so they can't be grouped into months or years.
- `D2` (country) is blank spaces, and `cicago` is misspelt. These show up as odd pivot rows.
- Sales use `RANDBETWEEN`, so after any recalculation the pivot is **stale** until you hit Refresh.

> [EXAM] **Self-test:** *How do you get total profit by state?* → New pivot, State in Rows, Profit in Values (set to Sum).

---

### Sheet11: Excel Table, Two Pivots, OFFSET

**Data:** zone, customer type (`online`/`cash`/`credit card`), revenue → converted into **`Table3`** (select data → `Ctrl+T`).

**Why use a Table?**
- Auto-expands when you add rows, so pivots and formulas based on it grow too.
- Banded rows, filter buttons, and **structured references** like `Table3[revenue]`.

**Two pivots:** *Row = customer type, Values = Sum of revenue*, one built on the Table and one on a plain range.
[MISTAKE] Their grand totals differ (**29,817 vs 30,962**). That's a classic sign one pivot is **stale** or built on a different snapshot. Refresh both (`Alt+F5`).

**`H2: =OFFSET($A$2,0,0,COUNTA($A:$A),COUNTA($1:$1))`** is a **dynamic range** that grows with the data.

`OFFSET(start_cell, rows_down, cols_right, height, width)`
- Start at `A2`, move 0 rows and 0 columns.
- Height = number of items in column A. Width = number of items in row 1 (3).

[MISTAKE] `COUNTA($A:$A)` counts the **header too** (13), but you start at `A2`, so the range is one row too tall. That's why row 14 shows `0, 0, 0`. Fix:

```excel
=OFFSET($A$2,0,0,COUNTA($A:$A)-1,COUNTA($1:$1))
```

> [EXAM] **Self-test:** *Why use a table rather than a plain range?* → It grows automatically and keeps formulas/pivots/charts up to date.

---

### Sheet12: Pivot with Two Row Fields

**Data (`A3:C12`):** city, date, sales. The pivot has **City → Date** as nested row labels and **Sum of SALES** as values (Grand Total = 38,957). The title "PIVOT TABLE" is merged across `A1:D1`.

[MISTAKE] Dates are text (`10.04.2026`), so they appear as individual items instead of grouping by Year/Month. With real dates you could right-click → **Group** → Months/Years.

> [EXAM] **Self-test:** *How do you show only one city's data?* → Drag City to **Filters**, or use the row-label filter drop-down.

---

### Sheet13: SUM / AVERAGE and the Insert Tab

- `F2: =SUM(B2:E2)` is each student's total.
- `G2: =AVERAGE(B2:E2)` is each student's average.
- An embedded **screenshot** shows the same functions in vertical and horizontal layouts (total 344, average 57.33).

**Insert tab tour (visible in your screenshot):** PivotTable, Recommended PivotTables, Table, Pictures, Shapes, Icons, 3D Models, SmartArt, Screenshot, Recommended Charts, Maps, PivotChart, Sparklines (Line / Column / Win-Loss), Slicer, Timeline, Link, Comment, Text Box, Header & Footer, Equation, Symbol.

> [EXAM] **Self-test:** *Which feature gives a tiny chart inside one cell?* → Sparklines. *Clickable filter buttons for a pivot?* → Slicer (or Timeline for dates).

---

### Sheet14: SEQUENCE and Map Chart [365]

**`A1: =SEQUENCE(10,10,50,-5)`** fills a 10 × 10 grid starting at **50**, stepping by **−5** across then down (50, 45, 40 … −445).

`SEQUENCE(rows, [columns], [start], [step])`

**Filled map chart:** countries in `A13:A20` with salary in `B13:B20`, charted with Insert → Maps → **Filled Map**.
[MISTAKE] Map charts need recognisable place names and an internet connection. `chaina`, `thai`, `swi` may fail to plot. Use `China`, `Thailand`, `Switzerland`.

> [EXAM] **Self-test:** *Make a list 1 to 20.* → `=SEQUENCE(20)`. *Even numbers 2 to 20?* → `=SEQUENCE(10,1,2,2)`

---

### Sheet15: DATEDIF, TODAY, and Wildcard SUMIF

**Data:** name, age, dept, position, salary, joining date, duration, "more than 5 years".

```excel
G: =DATEDIF(F2,TODAY(),"Y")&" years, "&DATEDIF(F2,TODAY(),"YM")&" months, "&DATEDIF(F2,TODAY(),"MD")&" days"
H: =IF(DATEDIF(F2,TODAY(),"Y")>=5,"Yes","No")
I: =SUMIF(C2:C11,"phy",E2:E11)
J: =SUMIF(D2:D11,"*teacher*",E2:E11)
```

**DATEDIF(start, end, unit)** is a "hidden" function (not in the function list, but it works):

| Unit | Gives |
|---|---|
| `"Y"` | complete years |
| `"M"` | complete months |
| `"D"` | days |
| `"YM"` | months **after** the full years |
| `"MD"` | days **after** the full months |
| `"YD"` | days after the full years |

`&` joins text pieces. `TODAY()` returns today's date and updates daily.

**Wildcards:** `*` = any number of characters, `?` = exactly one, `~` = escape. `"*teacher*"` matches `"teacher "` even with the trailing space. (An exact `"teacher"` would **fail** because the cell contains `"teacher "`. Good catch to remember.)

[MISTAKE] In column I, ranges drift (`C3:C12`, `C4:C13`…). Lock them: `$C$2:$C$11`.

> [EXAM] **Self-test:** *Years of service only?* → `=DATEDIF(F2,TODAY(),"Y")`. *Flag staff with 10+ years?* → `=IF(DATEDIF(F2,TODAY(),"Y")>=10,"Yes","No")`

---

### Sheet16: Custom Function `AddNameSpaces` (VBA)

Input like `VidyaChaudhari` → output `Vidya Chaudhari`. Column C calls `=AddNameSpaces(B2)`.

[MISTAKE] **Important:** the macro code is **not stored in this file** (a plain `.xlsx` can't hold macros), so the formulas will show `#NAME?` if you reopen the file elsewhere. Macros require **`.xlsm`**.

**How to recreate it:** `Alt+F11` → Insert → Module → paste → save as `.xlsm`.

```vba
Function AddNameSpaces(txt As String) As String
    Dim i As Long, ch As String, result As String
    result = Left(txt, 1)
    For i = 2 To Len(txt)
        ch = Mid(txt, i, 1)
        If ch Like "[A-Z]" And Mid(txt, i - 1, 1) Like "[a-z]" Then
            result = result & " "
        End If
        result = result & ch
    Next i
    AddNameSpaces = result
End Function
```

**No-VBA alternative:** type the first result by hand, then press **`Ctrl+E` (Flash Fill)**. 365 users can also use `=TRIM(REGEXREPLACE(B2,"([a-z])([A-Z])","$1 $2"))`.

---

### Sheet17: Splitting Text and Numbers (`GETNAME`, `GetNumbers`)

Input `123vidya` → name `vidya` (column B via `=GETNAME(A1)`) and number `123` (column C via `=GetNumbers(A1)`). Same situation as Sheet16: they were user-defined VBA functions and need an `.xlsm` file.

```vba
Function GetNumbers(txt As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(txt)
        If Mid(txt, i, 1) Like "[0-9]" Then r = r & Mid(txt, i, 1)
    Next i
    GetNumbers = r
End Function

Function GetName(txt As String) As String
    Dim i As Long, r As String
    For i = 1 To Len(txt)
        If Not Mid(txt, i, 1) Like "[0-9]" Then r = r & Mid(txt, i, 1)
    Next i
    GetName = r
End Function
```

**Formula-only versions (data is always "3 digits + name"):**
```excel
=--LEFT(A1,3)          ← number
=MID(A1,4,LEN(A1))     ← name
```
Flash Fill (`Ctrl+E`) also works for both.

> [EXAM] **Self-test:** *Difference between `LEFT`, `RIGHT`, `MID`?* → `LEFT(text,n)` takes n characters from the start, `RIGHT` from the end, `MID(text,start,n)` from the middle.

---

### Sheet18 + city sheets: Filtering and Splitting Data

**Data (`A1:F15`):** name, gender, city, state, region, segment. **AutoFilter** is on (`Ctrl+Shift+L`).

You filtered by city and copied each result to its own sheet: `ahemdabad`, `mubai`, `jaypur`, `delhi`.

**Steps:** filter city → select the visible cells → `Ctrl+C` → new sheet → `Ctrl+V` (filters copy only visible rows). A faster way: a pivot with City in Filters → Options → **Show Report Filter Pages**.

[MISTAKE] `north ` has a trailing space and `mubai`, `ahemdabad`, `maharast` are misspelt. These break filters and lookups. Clean with `=TRIM()`.

---

### Sheet23: Raw Data

State, category, netsale (`RANDBETWEEN(4000,8000)`). It has no analysis yet, so it's practice material for a pivot (State × Category), a chart, or `SUMIFS`.

---

### Sheet20: E-mail Hyperlinks

Pending amounts with an e-mail column where each address is a clickable `mailto:` link. Create with `Ctrl+K` → **E-mail Address**. Remove with right-click → *Remove Hyperlink*.

---

## 3. Cheat Sheet: All Functions at a Glance

| Function | Syntax | Use |
|---|---|---|
| `SUM / AVERAGE / MAX / MIN` | `=SUM(B2:E2)` | Basic stats |
| `IF` | `=IF(test, yes, no)` | Decisions |
| `AND / OR` | `=AND(a>1,b>1)` | Multiple conditions |
| `COUNT / COUNTA / COUNTBLANK` | `=COUNTA(A1:G12)` | Count numbers / non-empty / empty |
| `COUNTIF` | `=COUNTIF(range, crit)` | Count with 1 condition |
| `SUMIF` | `=SUMIF(range, crit, sum_range)` | Sum with 1 condition |
| `SUMIFS` | `=SUMIFS(sum_range, r1, c1, r2, c2)` | Sum with many conditions |
| `AVERAGEIF(S)` | `=AVERAGEIFS(avg, r1, c1, r2, c2)` | Average with conditions |
| `DSUM` | `=DSUM(db, field, criteria)` | Sum using a criteria block |
| `VLOOKUP` | `=VLOOKUP(val, table, col, 0)` | Look right |
| `XLOOKUP` [365] | `=XLOOKUP(val, look, return)` | Look any direction |
| `INDEX` | `=INDEX(range, row, col)` | Value at a position |
| `MATCH` | `=MATCH(val, range, 0)` | Position of a value |
| `UNIQUE` [365] | `=UNIQUE(range)` | Distinct values |
| `SEQUENCE` [365] | `=SEQUENCE(r, c, start, step)` | Number grid |
| `FREQUENCY` | `=FREQUENCY(data, bins)` | Count per bin |
| `PMT / IPMT / PPMT` | `=PMT(rate/12, n, -pv)` | Loan EMI / interest / principal |
| `DATEDIF` | `=DATEDIF(start, end, "Y")` | Date difference |
| `TODAY` | `=TODAY()` | Today's date |
| `OFFSET` | `=OFFSET(ref, r, c, h, w)` | Dynamic range |
| `RANDBETWEEN` | `=RANDBETWEEN(1,100)` | Random integers |
| `LEFT / MID / RIGHT` | `=MID(text, start, n)` | Cut text |
| `TRIM` | `=TRIM(text)` | Remove extra spaces |

---

## 4. Step-by-Step How-Tos

**Drop-down list:** select cell → Data → Data Validation → Allow **List** → Source `=$B$2:$B$16` (or type items separated by commas).

**Conditional formatting:** select range → Home → Conditional Formatting → Highlight Cells Rules → choose rule → pick colour. Remove via *Clear Rules* or *Manage Rules*.

**Pivot table:** click in data → Insert → PivotTable → drag fields → right-click → Refresh when data changes.

**Pivot chart:** click in the pivot → Insert → PivotChart.

**Excel Table:** click in data → `Ctrl+T` → tick "My table has headers" → rename in Table Design.

**AutoFilter:** `Ctrl+Shift+L`.

**Freeze values:** `Ctrl+C` → `Ctrl+Alt+V` → Values → Enter.

**Flash Fill:** type one example → `Ctrl+E`.

**Hyperlink:** `Ctrl+K`.

**Merge cells:** Home → Merge & Center.

**Absolute reference:** click inside the reference → `F4` (cycles `$A$1` → `A$1` → `$A1` → `A1`).

---

## 5. Mock Exam: Try Without Looking, Then Check

**Q1.** Marks in `C2:E2`. Show "Pass" only if every subject ≥ 40.
**Q2.** Count how many employees have designation "clerk" in `C2:C21`.
**Q3.** Total salary (`D`) of "vp" who are "m" (married) in `E`.
**Q4.** Average sale (`E`) for device "iphone" in `A`.
**Q5.** A drop-down in `G1` lists names in `B2:B16`; show their salary (column `D`) in `H1`.
**Q6.** EMI for ₹2,00,000 at 9% over 12 months.
**Q7.** Years completed between a joining date in `F2` and today.
**Q8.** Make a 5 × 5 grid starting at 100 and going up by 10.
**Q9.** Total of `B2:B20` where `A` contains the word "pro" anywhere.
**Q10.** List distinct regions from `B2:B20`.
**Q11.** Which feature lets you filter a pivot with clickable buttons?
**Q12.** Why might a pivot show old numbers?

<br>

**Answers**
1. `=IF(AND(C2>=40,D2>=40,E2>=40),"Pass","Fail")`
2. `=COUNTIF($C$2:$C$21,"clerk")`
3. `=SUMIFS($D$2:$D$21,$C$2:$C$21,"vp",$E$2:$E$21,"m")`
4. `=AVERAGEIF($A$2:$A$20,"iphone",$E$2:$E$20)`
5. `=VLOOKUP(G1,$B$2:$D$16,3,0)` (or `=XLOOKUP(G1,$B$2:$B$16,$D$2:$D$16)`)
6. `=PMT(9%/12,12,-200000)`
7. `=DATEDIF(F2,TODAY(),"Y")`
8. `=SEQUENCE(5,5,100,10)`
9. `=SUMIF($A$2:$A$20,"*pro*",$B$2:$B$20)`
10. `=UNIQUE(B2:B20)`
11. Slicer (Timeline for dates)
12. It hasn't been refreshed. Right-click → Refresh, or `Alt+F5`.

---

## 6. Your Personal Mistake Checklist (fix these habits)

- [ ] Lock ranges with `$` before filling down (Sheets 5, 6, 7, 15)
- [ ] End every nested IF with an "else" (Sheet1)
- [ ] Use the correct row in each formula (Sheet1 `J5`)
- [ ] Keep text out of numeric columns (Sheet1 column D)
- [ ] `FREQUENCY(data, bins)`: data first (Sheet2)
- [ ] Use the drop-down cell in the formula instead of a fixed cell (Sheets 3, 7)
- [ ] Same `type` in PMT, IPMT, PPMT (Sheet8)
- [ ] Same-size ranges across a group of COUNTIFs (Sheet9)
- [ ] Use real dates, not `10.10.2026` text (Sheets 10, 12)
- [ ] Refresh pivots before reading numbers (Sheets 10, 11)
- [ ] Subtract 1 for the header in `OFFSET`/`COUNTA` (Sheet11)
- [ ] Clean data with `TRIM` and spell-check before `UNIQUE`, pivots, or filters
- [ ] Save as `.xlsm` if you use VBA functions (Sheets 16, 17)

---

## 7. Last 10-Minute Revision

1. Re-read **Section 0** (Golden Rules) and **Section 3** (Cheat Sheet).
2. Cover the formulas in Section 2 and rebuild three of them from memory: `SUMIFS` grid, `VLOOKUP` + drop-down, `PMT`.
3. Do the **Mock Exam** with a stopwatch (aim for 2 minutes per question).
4. Practise one pivot table and one conditional formatting rule from scratch.

Good luck!
