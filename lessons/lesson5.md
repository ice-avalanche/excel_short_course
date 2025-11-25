# **Lesson 5. Pivot Tables and Pivot Charts**

## **1. Purpose & Use Cases**

A **Pivot Table** is one of Excel’s most powerful analytical tools.
It allows you to **summarize, group, filter, aggregate, and analyze** large datasets without writing formulas.

Pivot Tables turn flat, raw data into structured insights.

### **What Pivot Tables Can Do**

| Capability             | Description                            | Example                       |
| ---------------------- | -------------------------------------- | ----------------------------- |
| **Summarize data**     | Sum, count, average, min/max           | Total sales per month         |
| **Group information**  | Group dates, numbers, categories       | Group by quarter or age range |
| **Filter & slice**     | Display only relevant subsets          | Show a specific region        |
| **Drill-down**         | Double-click to see underlying records | View all orders for March     |
| **Build fast reports** | Rearrange data by dragging fields      | Region → Product → Month      |

### **Typical Use Cases**

| Scenario          | Pivot Table Output                     |
| ----------------- | -------------------------------------- |
| Sales analysis    | Sales by region, product, month        |
| HR analytics      | Headcount by department, gender, grade |
| Finance reporting | Expenses by category, quarter          |
| Operations        | Orders by status, processing time      |
| Marketing         | Campaign performance by channel        |

## **2. Creating a Pivot Table**

### **Steps to Create a Pivot Table**

| Step | Action                                                          |
| ---- | --------------------------------------------------------------- |
| 1    | Select any cell inside your dataset                             |
| 2    | Insert → **PivotTable**                                         |
| 3    | Choose data range (Excel auto-detects)                          |
| 4    | Choose destination: new or existing sheet                       |
| 5    | Drag fields into **Rows**, **Columns**, **Values**, **Filters** |

Once created, you can change layout instantly by rearranging fields.

## **3. Pivot Table Fields Explained**

Pivot Tables rely on **four types of field placements**:

| Field Area  | Purpose                           | Typical Fields                |
| ----------- | --------------------------------- | ----------------------------- |
| **Rows**    | Categories displayed vertically   | Region, Product, Employee     |
| **Columns** | Categories displayed horizontally | Month, Status, Metric         |
| **Values**  | Calculations                      | Sum of Sales, Count of Orders |
| **Filters** | High-level report filters         | Date, Department, Category    |

### **How It Works**

* Dragging **Region** to Rows → groups by region
* Dragging **Month** to Columns → creates a matrix
* Dragging **Sales** to Values → aggregates numbers
* Dragging **Year** to Filters → allows switching periods

### **Multiple Fields in One Area**

You can layer fields:
**Rows** → Region → Product → SKU
This creates a multi-level hierarchy.

## **4. Value Calculations (Summaries)**

The **Values** area calculates metrics such as:

| Calculation               | Description                     |
| ------------------------- | ------------------------------- |
| **Sum**                   | Total of numeric field          |
| **Count**                 | Number of records               |
| **Average**               | Mean value                      |
| **Max / Min**             | Highest / lowest values         |
| **% of Row/Column Total** | Share within row or column      |
| **Running Total**         | Cumulative sum                  |
| **Show Values As**        | % difference, rank, index, etc. |

### **Changing Calculation Type**

Right-click any value → **Summarize Values By** or **Show Values As**.

## **5. Grouping Data**

Pivot Tables can automatically group:

* **Dates** by year, quarter, month
* **Numbers** into bins
* **Text** into custom groups

## **6. Sorting and Filtering in Pivot Tables**

Pivot Tables provide extensive sorting and filtering:

### **Sorting**

| Type             | Example                     |
| ---------------- | --------------------------- |
| A→Z / Z→A        | Product names               |
| Smallest→Largest | Units sold                  |
| Custom Lists     | Mon, Tue, Wed…              |
| Top 10           | Top 10 customers by revenue |

### **Filtering**

| Filter       | Example                      |
| ------------ | ---------------------------- |
| Label Filter | “Contains ‘North’”           |
| Value Filter | “Sales > 10,000”             |
| Date Filter  | “This Month”, “Next Quarter” |

### **Report Filters**

Add fields to the **Filters** area:

* Year
* Department
* Category

User selects the value from a drop-down.

## **7. Pivot Table Formatting**

Formatting ensures readability and professionalism.

### **Key Formatting Tools**

| Tool                  | Path                              | Use                         |
| --------------------- | --------------------------------- | --------------------------- |
| **PivotTable Styles** | PivotTable Design → Styles        | Predefined themes           |
| **Number Format**     | Right-click value → Number Format | Currency, %, comma          |
| **Grand Totals**      | PivotTable Design → Grand Totals  | On/Off for rows/columns     |
| **Subtotals**         | PivotTable Design → Subtotals     | Top/Bottom or Off           |
| **Report Layout**     | Compact / Outline / Tabular       | Multi-level hierarchy style |

### **Common Layouts**

| Layout           | Description                            |
| ---------------- | -------------------------------------- |
| **Compact Form** | Saves space, nested rows in one column |
| **Outline Form** | Each field in its own column           |
| **Tabular Form** | Best for exporting, structured tables  |

## **8. Pivot Charts (PivotChart)**

A **PivotChart** visualizes Pivot Table summaries interactively.

### **Advantages**

| Benefit              | Description                              |
| -------------------- | ---------------------------------------- |
| Automatic updates    | Changes in Pivot Table reflect in chart  |
| Interactive          | Slicers & filters apply to chart         |
| Summarized data      | No need to prepare chart-specific ranges |
| Multiple chart types | Column, line, area, bar, combo           |

### **Creating a PivotChart**

1. Select Pivot Table
2. Insert → **PivotChart**
3. Choose chart type
4. Customize using **Chart Design** and **Format** tabs

### **Best Chart Types for Pivot Data**

| Chart      | Use Case                                 |
| ---------- | ---------------------------------------- |
| Column/Bar | Category comparison                      |
| Line       | Trends over time                         |
| Area       | Cumulative metrics                       |
| Combo      | Dual metrics (e.g., Sales vs Margin)     |
| Pie        | Composition (small number of categories) |

## **9. Slicers & Timelines**

**Slicers** and **Timelines** are interactive filters that control Pivot Tables and PivotCharts.

### **Slicers**

Visual buttons for filtering categorical data.

| Field   | Slicer Example              |
| ------- | --------------------------- |
| Region  | North / South / East / West |
| Product | Laptop / Monitor / PC       |
| Status  | Open / Closed / Pending     |

Path: Insert → **Slicer**

### **Timeline**

Dedicated filter for **dates**.

| Timeline Mode | Example         |
| ------------- | --------------- |
| Year          | 2023            |
| Quarter       | Q1, Q2          |
| Month         | Jan, Feb        |
| Day           | Individual days |

Path: Insert → **Timeline**

### **Advantages**

* Clear visual filtering
* Combine multiple slicers
* Control several Pivot Tables at once
* Excellent for dashboards

### **Connecting a Slicer to Multiple PivotTables**

Right-click slicer → **Report Connections** → check PivotTables.
