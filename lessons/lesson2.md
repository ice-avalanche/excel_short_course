# Lesson 2. Data Processing and Analysis

## Data Preparation & Cleaning (Tools & Functions)

| Tool / Function        | Description                                                                                                                         | Example / Path                                                                                                     |
| ---------------------- | ----------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ |
| **Text to Columns**    | Splits text from one column into multiple columns by delimiter or fixed width.                                                      | **Data → Text to Columns**. Example: split `Ivanov Ivan Ivanovich` into *Last*, *First*, *Middle*.                 |
| **Flash Fill**         | Detects a pattern from user input and auto-fills the rest. Useful for merging/splitting text, extracting numbers/text, reformatting.| **Data → Flash Fill** (or `Ctrl+E`). Example: enter `Ivanov I.I.` → Excel generates initials from full names.      |
| **Remove Duplicates**  | Removes duplicate rows based on selected columns.                                                                                   | **Data → Remove Duplicates**. Example: de-duplicate a product list by `Product ID`.                                |
| **Find & Replace**     | Finds text/values and replaces them (supports wildcards).                                                                           | `Ctrl+H`. Example: replace `,` with `.` in numeric text.                                                           |
| **Data Validation**    | Restricts input: whole/decimal numbers, list, date, time, text length, or custom formula.                                           | **Data → Data Validation**. Examples: Whole number 1–100; List: `Open,Closed`; Custom: `=ISNUMBER(A1)`.            |
| `TRIM`                 | Removes extra spaces (keeps single spaces between words).                                                                           | `=TRIM("  Ivan  Ivanov  ")` → `"Ivan Ivanov"`                                                                     |
| `CLEAN`                | Removes non-printing characters.                                                                                                    | `=CLEAN(A2)`                                                                                                       |
| `VALUE`                | Converts text to number.                                                                                                           | `=VALUE("123,45")` (with proper locale/separators)                                                                 |
| `TEXT`                 | Formats a value with a pattern.                                                                                                    | `=TEXT(TODAY(),"yyyy-mm-dd")`                                                                                      |
| `DATE`                 | Builds a date from Y/M/D numbers.                                                                                                  | `=DATE(2025,9,16)`                                                                                                 |
| `TIME`                 | Builds a time from H/M/S numbers.                                                                                                  | `=TIME(10,30,0)`                                                                                                   |
| `LEFT / MID / RIGHT`   | Extracts substrings (alternative to Flash Fill).                                                                                    | `=LEFT("RU-123-AB",2)` → `"RU"`                                                                                    |
| `SUBSTITUTE`           | Replaces all occurrences of a substring.                                                                                           | `=SUBSTITUTE("A,B,C",",",";")` → `"A;B;C"`                                                                         |
| **Numbers → Dates**    | Convert text like `20210304` to a real date.                                                                                        | `=DATE(LEFT(A2,4), MID(A2,5,2), RIGHT(A2,2))` → 04.03.2021                                                         |
| **Reformat date**      | Change visual format without altering value.                                                                                        | `=TEXT(TODAY(),"yyyy-mm-dd")` → `"2025-09-16"`                                                                     |
| **Extract date parts** | Get year, month, day for grouping.                                                                                                 | `=YEAR(A2)`, `=MONTH(A2)`, `=DAY(A2)`                                                                              |

## What-If Analysis Tools

| Tool / Command             | Description                                                                                                                            | Example / Path                                                                                                                     |
| -------------------------- | -------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| **Consolidate**            | Aggregates several ranges (even across sheets/workbooks) into a summary using Sum, Average, Min, Max, etc. Can create links to source. | **Data → Consolidate**. Example: combine 3 grade tables (*Name, Score1, Score2, Score3*) → **Sum** or **Average** by labels.       |
| **Scenario Manager**       | Saves named input sets (“scenarios”) for the same model cells and compares results in a report.                                        | **Data → What-If Analysis → Scenario Manager**. Example: “Excellent Student” vs “Good Student” → compare average grade.            |
| **One-Variable Data Table**| Evaluates one formula for many values of a single input cell.                                                                          | **Data → What-If Analysis → Data Table**. Example: Salary model `=Base*Coef + Base*Bonus`.                                         |
| **Two-Variable Data Table**| Evaluates one formula for all combinations of two inputs.                                                                              | Same path. Example: Row inputs = Bonuses `{2%,5%,7%,10%,12%}`, Column inputs = Coefs `{1.2,1.3,1.4,1.5}`.                         |
| **Goal Seek**              | Iteratively finds an input value that makes a formula return a specified result.                                                       | **Data → What-If Analysis → Goal Seek**. Example: solve `x³−2x²+x−5≈0` by setting formula cell **To value** `0` and changing `x`. |

## Data Analysis Functions

| Tool / Function            | Description                                                                                            | Example |
| -------------------------- | ------------------------------------------------------------------------------------------------------ | ------- |
| **Forecast Sheet**         | Builds a forecast chart and table with 95% confidence interval; detects or accepts manual seasonality. | **Data → Forecast Sheet**. Input: regular-interval dates + values. Confidence default 95%. |
| `FORECAST.ETS`             | Exponential smoothing forecast for a target date.                                                      | `=FORECAST.ETS(target_date, values, timeline)` |
| `FORECAST.ETS.CONFINT`     | Confidence interval for ETS forecast at a target date.                                                 | `=FORECAST.ETS.CONFINT(target_date, values, timeline, 0.95)` |
| `FORECAST.ETS.SEASONALITY` | Returns detected seasonality (period).                                                                 | `=FORECAST.ETS.SEASONALITY(values, timeline)` |
| `FORECAST.ETS.STAT`        | Returns stats (e.g., MASE, SMAPE) for ETS model.                                                       | `=FORECAST.ETS.STAT(values, timeline, 1)` |
| `FORECAST.LINEAR`          | Linear regression forecast for a target x.                                                             | `=FORECAST.LINEAR(x, known_y, known_x)` |
| `CORREL`                   | Pearson correlation coefficient between two arrays (range −1…+1).                                      | `=CORREL(A1:A12,B1:B12)` → relationship between *sales* and *ad spend*. |

## Validation – Typical Rules

| Need               | Validation Rule     | Example                                                 |
| ------------------ | ------------------- | ------------------------------------------------------- |
| Whole number range | Allow only 1…100    | **Data Validation → Whole number → between 1 and 100**  |
| Dropdown list      | Predefined statuses | **Allow: List** → `Open,In Progress,Closed`             |
| Date only          | Date in a range     | **Allow: Date** → between `01/01/2020` and `12/31/2030` |
| Custom numeric     | Only numbers in A2  | **Allow: Custom** → `=ISNUMBER(A2)`                     |
| Cross-cell rule    | A1 must be > B1     | **Allow: Custom** → `=A1>B1`                            |
