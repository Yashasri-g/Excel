# Excel

Excel is a spreadsheet tool used for:
- Data storage
- Data cleaning
- Analysis
- Reporting
- Dashboarding
For analysts, Excel = data manipulation engine

Workbook: Employee_Data.xlsx
Sheets: Sheet1, Sales_Data, Dashboard

Rows → Numbers (1,2,3…)
Columns → Letters (A,B,C…)

A cell reference : Column + Row
Example: B5

| Action                   | Shortcut            |
| ------------------------ | ------------------- |
| Move to last row of data | Ctrl + ↓            |
| Move to last column      | Ctrl + →            |
| Go to beginning          | Ctrl + Home         |
| Select entire row        | Shift + Space       |
| Select entire column     | Ctrl + Space        |
| Select full data table   | Ctrl + A            |
| Switch sheets            | Ctrl + Page Up/Down |

| Action        | Shortcut       |
| ------------- | -------------- |
| Copy          | Ctrl + C       |
| Cut           | Ctrl + X       |
| Paste         | Ctrl + V       |
| Paste Special | Ctrl + Alt + V |

# Freeze Panes
When working with large data:
View → Freeze Panes → Freeze Top Row
Used for:
- Headers visibility
- Large datasets

# Flash Fill
Auto-extract patterns (names, emails). -> Crtl + E 

Types of Cell Referencing

There are 3 types:

1️⃣ Relative Reference
2️⃣ Absolute Reference
3️⃣ Mixed Reference

1️⃣ RELATIVE REFERENCE (Default)
📌 Format:
A1
No $ sign.

📌 What happens?
When copied, the reference changes automatically.

✔ It adjusts automatically.

🧠 When to use: Row-by-row calculations , Standard arithmetic, Dynamic tables

2️⃣ ABSOLUTE REFERENCE
📌 Format:
$A$1
Dollar signs lock both: Column, Row
Example: Suppose:
Tax rate is in cell E1
You calculate:
=A2*$E$1
When dragged down:
A2 changes → A3, A4
$E$1 stays fixed
🧠 When to use:
Fixed tax rate
Constant discount
Fixed multiplier
Lookup reference

3️⃣ MIXED REFERENCE

Locks either row or column.
🔹 Lock Column Only
$A1
🔹 Lock Row Only
A$1
🔍 Why Mixed Reference Exists
Used in:
Multiplication tables
Matrix calculations
Advanced modeling

| Function    | Purpose         | Counts Text? | Multiple Conditions? |
| ----------- | --------------- | ------------ | -------------------- |
| `SUM()`     | Adds numbers    | ❌            | ❌                    |
| `AVERAGE()` | Calculates mean | ❌            | ❌                    |
| `MIN()`     | Smallest value  | ❌            | ❌                    |
| `MAX()`     | Largest value   | ❌            | ❌                    |

| Function     | What It Counts                     |
| ------------ | ---------------------------------- |
| `COUNT()`    | Numbers only                       |
| `COUNTA()`   | Non-empty cells                    |
| `COUNTIF()`  | Count based on 1 condition         |
| `COUNTIFS()` | Count based on multiple conditions |

| Function     | What It Counts                     |
| ------------ | ---------------------------------- |
| `COUNT()`    | Numbers only                       |
| `COUNTA()`   | Non-empty cells                    |
| `COUNTIF()`  | Count based on 1 condition         |
| `COUNTIFS()` | Count based on multiple conditions |

* Table is a structured dataset with built-in filtering, formatting, and dynamic range expansion.
* Pivot table is a tool used to summarize large datasets quickly without writing the formulas
* Tool used to summarize large datasets quickly.

| Chart        | When to Use           |
| ------------ | --------------------- |
| Column Chart | Compare categories    |
| Bar Chart    | Horizontal comparison |
| Line Chart   | Trends over time      |
| Pie Chart    | Percentage share      |
| Scatter Plot | Correlation analysis  |

- Chart directly connected to Pivot Table.
When Pivot changes → Chart updates automatically.
- Add interactive filters to Pivot Tables.
Insert → Slicer → Choose field
Used in dashboards heavily.

| Feature             | Normal Chart | Pivot Chart |
| ------------------- | ------------ | ----------- |
| Dynamic filtering   | ❌            | ✅           |
| Connected to pivot  | ❌            | ✅           |
| Best for dashboards | ❌            | ✅           |

--> Page Layout is mainly used when:
- Printing reports
- Exporting to PDF
- Submitting formatted reports

Margins - Control white space around sheet. 
Orientation - Portrait , Landscape (most reports use Landscape)
Size - A4 (commonly used in India)
Print Area - Select only specific area to print.
Scale to Fit - Shrink content to fit one page.
Page Break Preview - See where page splits

# BASIC DATA MANAGEMENT
1. Sorting
Arrange data:
A → Z
Z → A
Smallest → Largest
2. Filtering
Shows only required data.
Data → Filter → Select values
3. Remove Duplicates
Data → Remove Duplicates

* Consolidating Data - Combine data from multiple sheets.
* Template = Pre-designed Excel file used repeatedly.
* Dashboard = Visual summary of KPIs.

Basic Dashboard Components
* Pivot Table
* Pivot Chart
* Slicer
* KPI Cards
* Conditional formatting

Beginner Dashboard Steps
1. Clean Data
2. Convert to Table
3. Create Pivot Table
4. Insert Pivot Chart
5. Add Slicer
6. Design layout neatly

DATA VALIDATION - Controls what users can enter

Types:
- List (Dropdown)
Example:
Department: HR, IT, Sales
Data → Data Validation → List

- Whole Number - Allow only numbers between 1–100.

- Decimal - Limit salary range.

- Date - Allow only future dates.

- Custom Formula
Example:
Prevent duplicate Employee ID:
=COUNTIF(A:A,A1)=1

🎯 Why Analysts Use Data Validation
Prevent wrong entries
Maintain data integrity
Control user input
Avoid reporting errors

IF Function 
Returns one value if condition is TRUE, another if FALSE.
=IF(condition, value_if_true, value_if_false)

AND Function
Returns TRUE only if all conditions are TRUE
=AND(A2>30000, B2="IT")

OR Function
Returns TRUE if any condition is TRUE
=OR(A2="HR", A2="IT")

NOT Function
Reverses logic.
=NOT(A2>50000)

Nested IF
Using IF inside IF.
Example:
=IF(A2>50000,"High",
   IF(A2>30000,"Medium","Low"))

IFS Function (Cleaner)
=IFS(
A2>50000,"High",
A2>30000,"Medium",
TRUE,"Low"
)
Why better? --> No multiple closing brackets, Easier to read

LOOKUP Functions

Purpose:
Retrieve data from another table.

Example:
Fetch salary band from master table.
Lookup always needs:
1. Lookup value
2. Lookup table
3. Return column

LOOKUP Functions
Main types:
VLOOKUP - Vertical lookup 
HLOOKUP - horizontal lookup 
XLOOKUP
INDEX + MATCH

1. VLOOKUP
=VLOOKUP(A2, TableRange, ColumnNumber, FALSE)
2. INDEX Function

TEXT Functions - Used for cleaning and extracting text.
LEFT()
RIGHT()
MID()
LEN()
FIND()
TRIM()
UPPER()
LOWER()
PROPER()

TODAY() - Returns current date.
NOW() - Returns current date & time.
DATE() - Create custom date.
YEAR(), MONTH(), DAY() - Extract components.
DATEDIF() - Calculate difference between two dates.
EOMONTH() - Get end of month.

Sparklines are mini charts inside a single cell. Used to show trend visually.

| Error   | Meaning             |
| ------- | ------------------- |
| #DIV/0! | Division by zero    |
| #N/A    | Lookup not found    |
| #VALUE! | Wrong data type     |
| #REF!   | Invalid reference   |
| #NAME?  | Wrong function name |
** Error handling 
1. IFERROR Function
Prevents ugly errors.
=IFERROR(A2/B2,0)
=IFERROR(VLOOKUP(...),"Not Found")
2. ISERROR / ISNA
Check specific error types.
=IF(ISNA(VLOOKUP(...)),"Missing","Found")

Why Error Handling is Important
- Makes dashboards clean
- Avoids breaking models
- Professional presentation
- Prevents confusion in reports

Power Query = Data cleaning & transformation tool.
Think of it as: Python-like data preprocessing inside Excel (no coding required).
🔹 What It Can Do
Import data (Excel, CSV, PDF, Web, Folder)
Remove duplicates
Split columns
Merge files
Pivot / Unpivot
Change data types
Clean messy data
Append multiple sheets

Key Concepts
Query Editor
Applied Steps
M Language (behind the scenes)
Refresh Model

Power Pivot
Power Pivot allows:
Large datasets (millions of rows)
Relationships between tables
Data modeling
DAX calculations
Think: Excel + Database logic.

DAX FUNCTIONS - Used inside Power Pivot.
Think of DAX as: Advanced formula language for data models.
SUM, CALCULATE, COUNTROWS, fILTER

What is a Macro?
Recorded automation.
It uses VBA (Visual Basic for Applications).

🔹 What Macros Can Do
Automate repetitive tasks
Generate reports
Format sheets
Send emails
Create buttons
