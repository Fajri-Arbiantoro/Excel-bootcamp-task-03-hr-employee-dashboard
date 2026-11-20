# Project 03: Building an HR Employee Dashboard Using Advanced Excel

My third Excel Bootcamp project, documenting my progress in using advanced Excel features (lookups, Pivot Tables, Slicers, and VBA) to turn a raw employee sheet into an interactive HR dashboard.

## Project Background

This project was completed as an advanced personal task during an Excel Bootcamp I participated in.

The assignment involved working with a 123-row employee dataset and ten business questions about headcount, retirement age, KPI performance, salary, bonus, gender, and hiring history. The objective was to build helper columns with lookup formulas, summarize the data with a Pivot Table helper sheet, record and edit a VBA macro, and assemble everything into a one-click dashboard.

Compared with my first two Excel projects, which focused on cleaning sales records and linking student records across worksheets, this assignment gave me my first hands-on experience with Pivot Charts, Slicers, and VBA.

This project represents another step in developing practical data analysis skills using Microsoft Excel.

## Project Files

- [Anonymized Excel Workbook (.xlsm, macro-enabled)](./Employee-Dashboard-Anonymized.xlsm)

## Learning Objectives

Through this assignment, I aimed to:

- Pull values from master sheets using INDEX + MATCH and VLOOKUP instead of typing them in.
- Calculate age, flag retirement-age employees, and classify KPI performance with logical functions.
- Build a salary and bonus engine driven by Master tables.
- Summarize the data with Pivot Tables and Pivot Charts.
- Create an interactive employee card using Data Validation.
- Filter the whole dashboard with Slicers.
- Record and edit a simple VBA macro to clean data.

## Tasks & Analysis Process

The assignment consisted of eight main steps:

**1. Prepare the Data with Lookups**

Helper columns pull Division and Gender from master sheets, and the hire year is extracted from the start date.

Functions used: INDEX, MATCH, YEAR.

**2. Add Business Logic: Age, Retirement and KPI**

Calculated age at a reference date, flagged employees reaching retirement age (50), and classified KPI as Good or Poor (Poor = below 70%). The reference date and retirement age live in assumption cells, so no value is hard-coded inside a formula.

Functions used: YEARFRAC, INT, IF.

**3. Build the Salary & Bonus Engine**

Looked up each employee's position, salary, and bonus multiplier from the Master tables. Bonus is paid only when KPI is above 90%.

Functions used: VLOOKUP, IF.

**4. Create the Pivot Table Helper Sheet**

Built six Pivot Tables that answer the business questions and feed every chart on the dashboard.

**5. Build an Interactive Employee Card**

Used a Data Validation dropdown to select an employee name, then INDEX + MATCH to return that person's Base, Employee Type, KPI, and Salary.

**6. Add Slicers**

Added three Slicers (Employee Type, Position, and Base) and connected them to the Pivot Charts through Report Connections.

**7. Assemble the Dashboard**

Combined KPI and salary cards, a gender donut chart, a historical hiring combo chart (bars for hires per year, line for cumulative headcount), a total-by-base chart, and a gender-per-division chart.

**8. Automate with VBA**

Recorded a macro that deletes rows with a blank employee name, then edited it to show a confirmation message box when the cleaning is finished.

## Key Findings

Based on the completed workbook:

1. Permanent employees based in Jakarta: 75.
2. Employees reaching retirement age (50) by the end of 2025: 3.
3. Employees with a "Poor" KPI (below 70%): 32, or about 26% of 123.
4. Salary contribution by position: General Manager 25.80%, Supervisor 18.39%, Senior Manager 16.89%, Senior Staff 16.42%, Staff 12.01%, Junior Manager 6.75%, CEO 3.75%.
5. Total bonus paid: Rp201,000,000 (Senior Manager Rp108,000,000, Supervisor Rp49,000,000, Senior Staff Rp40,000,000, Staff Rp4,000,000).
6. The workforce is about 54.5% male and 45.5% female.
7. Hiring peaked in 2018 with 17 hires, and 109 of the 123 employees are based in Jakarta.

## Tools & Techniques

**Tool:** Microsoft Excel (with VBA)

**Skills Practised:**
- Lookup Functions (INDEX + MATCH, VLOOKUP)
- Logical Functions
- Date Calculations
- Assumption Cells (no hard-coding)
- Pivot Tables and Pivot Charts
- Data Validation
- Slicers
- Dashboard Design
- VBA Macro Recording and Editing

## Challenges & Learning Reflection

The biggest challenge in this project was VBA. It was completely new to me, and understanding objects such as Selection and SpecialCells, and how Excel records actions, took the most time. It is also the area I found most interesting.

Another important takeaway was building the workbook so it can be reused. Keeping the reference date and retirement age in cells, and pulling division, gender, position, salary, and bonus rules from master tables, means a single change updates every row. This made the model much more flexible than typing values directly into formulas.

I also saw how Pivot Tables, Slicers, and charts work together as a pipeline: clean data, helper columns, pivots, charts, slicers, and finally a dashboard that answers HR questions in one click.

## What's Next?

This project builds on the lookup, conditional logic, and summary techniques I practised in my first two Excel assignments.

I plan to continue strengthening my Excel skills, especially VBA, and gradually expand my toolkit with SQL, Python, and data visualization tools as I work toward a career in data analytics.

I will continue documenting each project, including the techniques I practise, the challenges I encounter, and how my analytical skills develop over time.

---

**Disclaimer**

This project was completed as part of an advanced Excel Bootcamp for educational and portfolio purposes only. The analysis is based on an employee practice dataset supplied for the assignment. The workbook shared in this repository is an anonymized version prepared for portfolio demonstration: employee names and birthplaces were replaced with generic labels, birth dates were changed to different days and months within the same year (so ages and results are unchanged), and file metadata was removed. The original, non-anonymized workbook is not shared because it may contain personal information. Please do not attempt to identify individuals or use the dataset for purposes beyond this educational demonstration. The project is not intended for commercial use or to represent actual employee data.
