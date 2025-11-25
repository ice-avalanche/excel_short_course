# **Lesson 6. Power Query. Import and Data Preparation**

## **1. Introduction to Power Query**

**Power Query** is Excel’s built-in ETL engine: **Extract → Transform → Load**.
It allows you to import data from almost any source, clean it, shape it, and refresh it automatically without writing formulas.

### **What Power Query Does**

| Stage         | Description                      | Example                          |
| ------------- | -------------------------------- | -------------------------------- |
| **Extract**   | Connect to various sources       | Excel, CSV, SQL, Web             |
| **Transform** | Clean & reshape data             | Remove duplicates, split columns |
| **Load**      | Output into Excel                | Table, Data Model                |
| **Refresh**   | Re-apply all steps automatically | Weekly/monthly reporting         |

### **Where Power Query Lives**

Excel Ribbon → **Data → Get & Transform Data**

### **Why It Matters**

* Automates repetitive tasks
* Ensures consistent results
* Handles large datasets
* Integrates with Power BI
* Ideal for scalable reporting workflows

## **2. Importing Data**

Power Query supports **dozens of connectors**.

### **Common Import Sources**

| Category      | Examples                       |
| ------------- | ------------------------------ |
| **Files**     | Excel, CSV, TXT, XML, JSON     |
| **Folders**   | Combine multiple monthly files |
| **Databases** | SQL Server, Oracle, MySQL      |
| **Cloud**     | SharePoint, OneDrive           |
| **Web / API** | HTML tables, JSON endpoints    |

### **File Import Example (CSV)**

Steps:

1. Data → **Get Data → From File → From Text/CSV**
2. Preview window opens
3. Configure delimiter & encoding
4. Load or Transform

Power Query automatically detects:

* column headers
* data types
* delimiters

### **Folder Import (Combining Files)**

Used for:

* monthly reports
* daily logs
* multi-file data dumps

Power Query:

* reads folder
* filters files
* combines them automatically
* applies same transformations to each file

## **3. Power Query Editor Overview**

The **Editor** is where all transformations happen.

### **Main Components**

| Component                | Description                             |
| ------------------------ | --------------------------------------- |
| **Ribbon**               | Access to all transformation commands   |
| **Queries Pane**         | List of all queries                     |
| **Preview Grid**         | Data preview (not the full dataset)     |
| **Query Settings**       | Applied steps, naming                   |
| **Formula Bar**          | Underlying M code                       |
| **Data Profiling Tools** | Column quality, distribution, profiling |

### **Applied Steps**

Every transformation becomes a step:

* promotes headers
* removes columns
* filters rows
* splits text
* merges queries

Steps are fully reproducible on refresh.

## **4. Data Transformation Tools**

Power Query includes an extensive set of tools for shaping data.

### **Column Operations**

* Remove / keep columns
* Split by delimiter
* Merge columns
* Change data types
* Duplicate & reference
* Add custom columns

### **Row Operations**

* Remove top/bottom rows
* Remove duplicates
* Filter rows
* Sort ascending/descending

### **Text Transformations**

* Trim
* Clean
* Uppercase/lowercase
* Extract before/after delimiter
* Replace values

### **Number Transformations**

* Round
* Multiply/divide
* Absolute value
* Handling errors

### **Date & Time Tools**

* Extract year/month/quarter
* Age calculation
* Combine date/time

### **Structural Tools**

* Pivot
* Unpivot
* Group By
* Transpose

These tools enable full-scale data cleaning and reshaping.

## **5. Merging and Appending Queries**

Two essential ways to combine datasets:

### **Merge Queries (JOIN)**

Combines tables **horizontally**.

| Join Type   | Purpose              |
| ----------- | -------------------- |
| Left Outer  | Most common lookup   |
| Right Outer | Reverse lookup       |
| Inner       | Only matching rows   |
| Full Outer  | All rows from both   |
| Anti Joins  | Find missing matches |

Example:
Merge Sales with Products on ProductID.

### **Append Queries (UNION)**

Stacks tables **vertically**.

Used for:

* monthly files
* yearly archives
* similar structures from different systems

Example:
Append Sales_2024 + Sales_2025.

## **6. Data Types (Critical Concept)**

Data types are **fundamental** to Power Query.
Incorrect types cause:

* merge mismatches
* calculation errors
* sorting/filtering issues
* refresh failures

### **Key Types**

| Icon       | Type         | Use                |
| ---------- | ------------ | ------------------ |
| ABC        | Text         | Names, IDs         |
| 123        | Whole Number | Counts, quantities |
| 1.2        | Decimal      | Prices, rates      |
| Calendar   | Date         | Transaction dates  |
| Clock      | Time         | Shifts, logs       |
| Date/Time  | Timestamp    | System logs        |
| True/False | Logical      | Yes/No             |

### **Important**

* Join keys *must* match data types
* Date columns must be real dates, not text
* Locale settings matter for date/decimal parsing

Always fix data types early in the transformation chain.

## **7. Complex Transformations**

Advanced shaping techniques include:

### **Pivot & Unpivot**

* Pivot → turn values into columns
* Unpivot → essential for cleaning monthly “wide” files

### **Group By**

Summaries:

* Sum
* Average
* Min/Max
* Nested tables

### **Conditional Columns**

Logic examples:

* “If Sales > 1000 then ‘High’ else ‘Low’”
* Category mapping
* Threshold classification

### **Fill Down / Fill Up**

Fixes hierarchical or merged cell data.

### **Working with Nested Structures**

* Lists
* Records
* Nested tables (from merges or group-by)

These enable professional ETL pipelines.

## **8. M-Code Basics**

M is the language behind Power Query.

### **Basic Structure**

```m
let
    Source = ...,
    Step1 = ...,
    Step2 = ...
in
    Step2
```

### **Common M Functions**

| Type    | Examples                          |
| ------- | --------------------------------- |
| Text    | Text.Upper, Text.Split            |
| Number  | Number.Round, Number.Abs          |
| Date    | Date.Year, Date.Month             |
| Table   | Table.SelectRows, Table.AddColumn |
| List    | List.Sum, List.Count              |
| Logical | if … then … else                  |

### **Error Handling**

```m
try Number.From([Value]) otherwise null
```

### **Custom Logic Example**

```m
= Table.AddColumn(Source, "Total", each [Units] * [Price])
```

Understanding M code makes complex transformations easier and more reliable.

## **9. Loading Data**

Power Query can load to:

### **1. Excel Table (Worksheet)**

* Visible
* Easy to inspect
* Limited by Excel row count
* Slower with large data

### **2. Data Model (Highly Recommended)**

* Efficient compression
* Large-scale analytics
* Supports relationships & DAX
* Ideal for dashboards

### **3. Connection Only**

* For staging queries
* Keeps workbook clean
* Improves performance

### **4. PivotTable / PivotChart**

* Directly builds dashboards
* Automatically refreshes

### **Loading Strategy**

| Scenario                   | Recommended Load         |
| -------------------------- | ------------------------ |
| Small dataset              | Worksheet                |
| Large dataset (>500k rows) | Data Model               |
| Multi-step ETL             | Connection + Data Model  |
| Dashboards                 | Data Model + PivotTables |

## **10. Refresh & Automation**

Refresh is the automation engine of Power Query.

### **What Refresh Does**

* Re-imports data
* Re-applies all steps
* Reloads outputs
* Updates PivotTables & charts

### **Refresh Methods**

| Method                  | Use Case                        |
| ----------------------- | ------------------------------- |
| Refresh All             | Full reporting update           |
| Refresh Query           | Update single pipeline          |
| Refresh on Open         | Automated dashboards            |
| Refresh Every X Minutes | Monitoring/logs                 |
| Background Refresh      | Continue working during refresh |

### **Automation Scenarios**

* Monthly sales consolidation
* Weekly HR reports
* Finance reconciliation pipelines
* Daily operational dashboards
* Folder-based auto file ingestion
