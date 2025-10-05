# Lesson 3. Excel Tables, Conditional Formatting, Filters & Sorting

## 1. Excel Tables (Smart Tables)

**Excel Tables** are structured data ranges that automatically expand, format, and support **named references** instead of static cell coordinates.  
They simplify analysis, ensure consistency, and integrate seamlessly with formulas, charts, and Power Query.

**Benefits:**
- Automatic range expansion when adding new rows or columns  
- Instant formula propagation  
- Built-in filters and sorting  
- Consistent styling  
- Easy integration with charts and PivotTables  

### Creating a Table

| Method | Steps |
|--------|-------|
| **Shortcut (Ctrl + T)** | 1. Select range with headers → 2. Press **Ctrl + T** → 3. Check *My table has headers* → 4. OK |
| **Home → Format as Table** | Select range → choose style → confirm headers |
| **Insert → Table** | Select range → confirm → check headers box |

Once created, Excel adds filters, formatting, and the **Table Design** tab for customization.

### Key Features

| Feature | Description | Example |
|----------|-------------|----------|
| **Auto-expansion** | Automatically includes new rows/columns | Adding a new entry expands the table |
| **Auto-fill formulas** | Formulas fill entire column automatically | `=[@Qty]*[@Price]` |
| **Totals Row** | Add via **Table Design → Total Row** | SUM, AVERAGE, COUNT, etc. |
| **Structured references** | Readable formulas using column names | `=SUM(Sales[Revenue])` |
| **Integration** | Tables sync with charts, PivotTables, and Power Query | Updates reflected automatically |

### Structured References

Instead of cell ranges (`B2:B10`), use column names:
```

=FUNCTION(TableName[ColumnName])

```

| Task | Formula | Explanation |
|------|----------|-------------|
| Total Revenue | `=SUM(Sales[Revenue])` | Adds up all revenue values |
| Average Salary | `=AVERAGE(Employees[Salary])` | Calculates mean |
| Row Calculation | `=[@Qty]*[@Price]` | Uses current row’s values |
| Multi-column range | `=AVERAGE(Sales[[Qty]:[Price]])` | Refers to range within table |

**Special Qualifiers:**
| Reference | Meaning |
|------------|----------|
| `[#All]` | Entire table |
| `[#Data]` | Only data rows |
| `[#Headers]` | Header row |
| `[#Totals]` | Totals row |

### Best Practices
* Always convert data to a Table before analysis.  
* Use descriptive table and column names.  
* Avoid merged cells inside tables.  
* Combine with Power Query or PivotTables for flexible analysis.  
* Verify automatic expansion after adding new data.  

## 2. Conditional Formatting

**Conditional Formatting** visually emphasizes trends and exceptions.  
It automatically changes cell appearance based on content or formulas — no manual edits needed.

**Path:** `Home → Conditional Formatting`

### Basic Rule Types

| Type | Description | Example |
|------|-------------|----------|
| Greater / Less Than / Between | Highlight numeric thresholds | `>100000` |
| Text Contains | Highlight keywords | “Manager” |
| Date Occurring | Highlight dates relative to today | “This Month” |
| Duplicate / Unique | Identify repeating or unique values | — |
| Top / Bottom N | Show best or worst performers | Top 10 sales |
| Above / Below Average | Compare against overall mean | “Above Average” values |

### Visual Styles

| Type | Description | Example |
|------|-------------|----------|
| **Data Bars** | In-cell horizontal bars show relative magnitude | Progress or performance |
| **Color Scales** | Gradient colors represent value range | Sales heatmap |
| **Icon Sets** | Arrows, circles, flags for value categories | KPI indicators |

### Formula-Based Rules

Use **“Use a formula to determine which cells to format”** for advanced logic.  
Formula returns **TRUE** for cells that need formatting.

| Goal | Formula | Explanation |
|------|----------|-------------|
| Highlight overdue tasks | `=$E2<TODAY()` | Checks date in column E |
| Highlight “Critical” rows | `=$F2="Critical"` | Locks column F |
| Mark <80% completion | `=$G2<0.8` | Numeric comparison |
| Stripe even rows | `=MOD(ROW(),2)=0` | Zebra pattern |

### Best Practices
* Use consistent color logic (red = risk, green = success).  
* Avoid too many overlapping rules.  
* Test formulas before applying.  
* Check active filters before exporting.  

## 3. Filters

**Filters** let you display only the rows that meet specific conditions.  
They’re ideal for quick data exploration without altering the dataset.

**Path:** `Data → Filter` or **Ctrl + Shift + L**

### AutoFilter

| Type | Description | Example |
|------|-------------|----------|
| Text | Filter by words or patterns | “Contains Manager” |
| Number | Filter by thresholds | “>100000” |
| Date | Filter by month, quarter, or year | “This Month” |
| Color | Filter by cell color or font color | Red = error |
| Top 10 | Show top/bottom values | “Top 10 Clients” |

### Advanced Filter

For complex logic or copying filtered results elsewhere.

| Step | Action |
|------|--------|
| 1 | Create criteria range with same headers. |
| 2 | Data → **Advanced Filter**. |
| 3 | Set list and criteria ranges. |
| 4 | Check *Unique records only* if needed. |

**Example (AND logic):**

| Department | Salary |
|-------------|---------|
| Sales | >100000 |

**Example (OR logic):**

| Department | Salary |
|-------------|---------|
| Sales |  |
|  | >100000 |

### Filtering by Color / Icon

If Conditional Formatting applied:
- Filter by **Cell Color**, **Font Color**, or **Icon**.  
Useful for dashboards and KPI views.

### Best Practices
* Include headers before filtering.  
* Clear all filters before analysis or export.  
* Avoid merged cells.  
* Use Excel Tables for automatically expanding filters.  

## 4. Sorting

**Sorting** organizes data logically — by values, text, date, or even color — to simplify comparison and reporting.

**Path:** `Data → Sort & Filter → Sort`

### Basic Sorting

| Data Type | Sort Order | Example |
|------------|-------------|----------|
| Numbers | Smallest → Largest / Largest → Smallest | Revenue by size |
| Text | A → Z / Z → A | Alphabetical order |
| Dates | Oldest → Newest / Newest → Oldest | Task deadlines |

* Always “Expand Selection” when sorting to avoid breaking row alignment.

### Multi-Level Sorting

Use **Custom Sort** for multiple criteria.

| Level | Column | Order |
|--------|--------|--------|
| 1 | Department | A → Z |
| 2 | Position | A → Z |
| 3 | Salary | Largest → Smallest |

**Path:** Data → Sort → Add Level → define sequence.

### Sorting by Color / Icon

If Conditional Formatting is applied:
- Sort by **Cell Color**, **Font Color**, or **Icon**.  
Example: Move red-highlighted errors to the top.

### Custom Lists

Create a custom sequence for logical ordering (e.g., weekdays, priorities).

**Path:** Data → Sort → Order → Custom List → enter:
```

Critical, High, Medium, Low

```

### Best Practices
* Convert data to Tables before sorting.  
* Sort entire dataset, not single columns.  
* Use consistent sort logic across reports.  
* Save frequently used orders as Custom Lists.  
