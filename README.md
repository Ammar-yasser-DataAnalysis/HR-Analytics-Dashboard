# 👥 HR Analytics Dashboard — Excel Only

> A fully interactive, dynamic HR analytics dashboard built **entirely in Microsoft Excel** — no Power BI, no Tableau. Powered by Power Query (Advanced Editor), Pivot Tables, and native Excel slicers for real-time cross-filtering across all visuals.

---

## 📌 Project Overview

This project transforms raw employee data into a rich, interactive HR dashboard using **only Excel's built-in capabilities**. The goal was to prove that professional-grade analytics dashboards don't require external BI tools — just a deep mastery of Excel's data pipeline.

---

## 🛠️ Tech Stack

| Layer | Tool |
|---|---|
| **Data Source** | Excel Workbook (`.xlsx`) |
| **ETL & Transformation** | Power Query (M Language — Advanced Editor) |
| **Data Modeling** | Excel Data Model (Pivot Cache) |
| **Aggregation** | Pivot Tables (multiple) |
| **Visualization** | Excel Charts (Bar, Line, Area) |
| **Interactivity** | Slicers connected to all Pivot Tables |

> ✅ Zero external tools — 100% native Excel

---

## ⚙️ Power Query Transformation (Advanced Editor)

All data cleaning and feature engineering was done in Power Query using M Language:

```m
let
    Source = Excel.Workbook(File.Contents("Employees.xlsx"), null, true),
    EmpTable_Table = Source{[Item="EmpTable",Kind="Table"]}[Data],

    // Step 1: Set correct data types
    #"Changed Type" = Table.TransformColumnTypes(EmpTable_Table,{
        {"No", Int64.Type}, {"First Name", type text}, {"Last Name", type text},
        {"Gender", type text}, {"Start Date", type date}, {"Years", Int64.Type},
        {"Department", type text}, {"Country", type text}, {"Center", type text},
        {"Monthly Salary", Int64.Type}, {"Annual Salary", Int64.Type},
        {"Job Rate", type number}, {"Sick Leaves", Int64.Type},
        {"Unpaid Leaves", Int64.Type}, {"Overtime Hours", Int64.Type}
    }),

    // Step 2: Merge first & last name
    #"Added Custom" = Table.AddColumn(#"Changed Type", "Full name",
        each [First Name] &" "& [Last Name]),

    // Step 3: Reorder columns
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{
        "No", "First Name", "Last Name", "Full name", "Gender", "Start Date",
        "Years", "Department", "Country", "Center", "Monthly Salary",
        "Annual Salary", "Job Rate", "Sick Leaves", "Unpaid Leaves", "Overtime Hours"
    }),

    // Step 4: Calculate Overtime Amount (1.5x hourly rate)
    #"Added Custom1" = Table.AddColumn(#"Reordered Columns", "Overtime amount",
        each (([Monthly Salary]/30/8)*1.5)*[Overtime Hours]),

    // Step 5: Calculate Daily Rate
    #"Added Custom2" = Table.AddColumn(#"Added Custom1", "Daily rate",
        each [Monthly Salary]/30),

    // Step 6: Calculate Net Salary (salary + overtime - unpaid leave deductions)
    #"Added Custom3" = Table.AddColumn(#"Added Custom2", "Net Salary",
        each ([Monthly Salary]+[Overtime amount])-([Daily rate]*[Unpaid Leaves])),

    // Step 7: Apply currency formatting
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Custom3",{
        {"Annual Salary", Currency.Type}, {"Monthly Salary", Currency.Type},
        {"Overtime amount", Currency.Type}, {"Daily rate", Currency.Type},
        {"Net Salary", Currency.Type}
    })
in
    #"Changed Type1"
```
---
<img src="https://github.com/Ammar-yasser-DataAnalysis/HR-Analytics-Dashboard/blob/main/Power%20query%20and%20advanced%20editor.png">

---

## 📐 Power Pivot & Data Model

The cleaned data was loaded directly into **Power Pivot** (Excel's built-in Data Model) where all calculated measures were defined — keeping the logic centralized and reusable across all Pivot Tables.

### Calculated Measures (DAX-style in Power Pivot)

| Measure | Description |
|---|---|
| `Numbers Of Employees` | Total headcount |
| `Numbers Of Department` | Distinct department count |
| `Male` | Count of male employees |
| `Female` | Count of female employees |
| `Excellent Performance` | Count of employees rated Excellent |
| `Medium Performance` | Count of employees rated Medium |
| `Not Bad Performance` | Count of employees rated Not Bad |
| `Bad Performance` | Count of employees rated Bad |
| `Fresh one` | Employee with minimum years of experience |
| `Oldest Employee` | Employee with maximum years of experience |
| `Best of job rate` | Maximum job rate value |
| `Bad Job rate` | Minimum job rate value |
| `AVG job rate` | Average job rate across all employees |
| `Avg Salary in month` | Average monthly salary |
| `avg Overtime Hours` | Average overtime hours per employee |
| `total Overtime Hours` | Sum of all overtime hours |
| `Total Monthly Salary` | Sum of all monthly salaries |
| `Total Annual Salary` | Sum of all annual salaries |

### Data Model Structure
- Single-table model (**EmpTable**) enriched with all engineered columns from Power Query
- Date hierarchy auto-generated: **Year → Quarter → Month**
- All Pivot Tables reference the same Data Model cache — ensuring consistency across the entire dashboard

### Engineered Columns

| Column | Formula Logic |
|---|---|
| `Full name` | `First Name & " " & Last Name` |
| `Overtime amount` | `(Monthly Salary / 30 / 8) × 1.5 × Overtime Hours` |
| `Daily rate` | `Monthly Salary / 30` |
| `Net Salary` | `(Monthly Salary + Overtime amount) − (Daily rate × Unpaid Leaves)` |

---

## 📊 Pivot Tables & Data Model

Multiple Pivot Tables were built on top of the Data Model to power each visual independently, all connected via shared slicers:

| Pivot Table | Metrics |
|---|---|
| Performance Summary | Employees by performance level (Excellent / Medium / Not Bad / Bad) |
| Gender Split | Male vs. Female headcount |
| Salary by Department | Total Monthly Salary & Average Net Salary per department |
| Performance by Country | Performance level distribution across Egypt, Lebanon, Saudi Arabia, Syria, UAE |
| Net Salary by Job Rate | Total Net Salary segmented by job rating |
| Salary by Country & Department | Drill-down net salary across country → department |
| Salary Timeline (Quarterly) | Net salary trend from 2016 Q1 to 2020 Q4 |
| Employees & Salary Over Time | Annual headcount vs. net salary (dual-axis) |
| Bad Performance by Department | Top departments with bad performance count |
| Salary by Year | Headcount and net salary summarized by year |

---

## 📈 Dashboard KPIs

| KPI | Value |
|---|---|
| Total Employees | **689** |
| Total Net Salary | **$1,504,793.39** |
| Average Net Salary | **$2,184.03** |
| Male Employees | **449** |
| Female Employees | **240** |
| Excellent Performance | **215** |
| Medium Performance | **332** |
| Not Bad Performance | **72** |
| Bad Performance | **70** |

---

## 🖥️ Dashboard Visuals

| Chart | Type | Insight |
|---|---|---|
| Total Salary per Performance Level | Horizontal Bar | Medium performers hold the highest total salary pool |
| Performance Levels by Country | Grouped Bar | Egypt dominates with 379 employees |
| Employees & Salaries Over Time | Dual-Axis Line | Peak in 2019 with 248 employees and $526K salary |
| Top 5 Bad Performance Departments | Column | Manufacturing leads with 16 bad performers |
| Salary Timeline (Quarterly) | Area Line | Quarterly salary trend 2016–2020 |
| Salary per Department | Horizontal Bar | Manufacturing tops at $280,649 |

---

## 🎛️ Interactivity — Dynamic Slicers

All slicers are connected to **all Pivot Tables simultaneously**, making the entire dashboard update in real-time on every filter selection:

- 🌍 **Country** — Egypt / Lebanon / Saudi Arabia / Syria / UAE
- 👤 **Gender** — Male / Female
- 📅 **Date/Year** — 2016 / 2017 / 2018 / 2019 / 2020
- ⭐ **Performance Level** — Bad / Medium / Not Bad / Excellent
- 🏢 **Department** — 18 departments

---
## Dashboard Screenshots (Click to enlarge) :
<img src="https://github.com/Ammar-yasser-DataAnalysis/HR-Analytics-Dashboard/blob/main/Dashboard.png">


---



## 🗃️ Dataset Overview

| Field | Description |
|---|---|
| `No` | Employee ID |
| `Full name` | Merged first + last name |
| `Gender` | Male / Female |
| `Start Date` | Employment start date |
| `Years` | Years of experience |
| `Department` | One of 18 departments |
| `Country` | Egypt, Lebanon, Saudi Arabia, Syria, UAE |
| `Monthly Salary` | Base monthly salary |
| `Annual Salary` | Annual base salary |
| `Job Rate` | Performance rating (numeric) |
| `Sick Leaves` | Number of sick leave days |
| `Unpaid Leaves` | Leave days deducted from salary |
| `Overtime Hours` | Extra hours worked |
| `Overtime amount` | *(Engineered)* Calculated overtime pay |
| `Daily rate` | *(Engineered)* Daily salary rate |
| `Net Salary` | *(Engineered)* Final take-home salary |

---

## 📁 Repository Structure

```
HR-Analytics-Excel-Dashboard/
│
├── README.md
├── Employees.xlsx              # Source data + Power Query + Data Model + Dashboard
└── Screenshots/
    ├── Dashboard.png           # Final interactive dashboard
    ├── Pivot_table_1.png       # KPI summary pivots
    ├── Pivot_table_2.png       # Country & salary pivots
    └── Pivot_table_3.png       # Performance & year pivots
```

---

## 🚀 How to Use

1. **Download** `Employees.xlsx`
2. **Open** in Microsoft Excel (2016 or later recommended)
3. Go to the **Dashboard** sheet
4. Use the **slicers** on the left panel to filter by Country, Gender, Year, Performance Level, or Department
5. All charts and KPI cards update **automatically** in real-time

> 💡 If data doesn't refresh, go to **Data → Refresh All**

---

## 💡 Key Insights

- **Manufacturing** is the highest-paying department ($280,649 total) and also has the most bad performers (16)
- **Egypt** accounts for **55%** of total headcount (379 out of 689)
- Salary peaked in **2019** at $526,435 with 248 employees, then declined in 2020
- **Medium performers** represent the largest group (332) but **Excellent performers** earn disproportionately more per head
- **Quality Control** and **Manufacturing** top the bad performance chart, suggesting training intervention opportunities

---

## 📄 License

This project is for educational and portfolio purposes.
