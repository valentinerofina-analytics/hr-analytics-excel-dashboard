# HR Analytics Excel Dashboard

An advanced HR Analytics Dashboard built using Python (openpyxl) and Microsoft Excel, analyzing 100 employee records across 7 departments.

## Project Overview

This project automates the generation of a professional HR Analytics Dashboard in Excel using Python. It demonstrates end-to-end data analysis — from raw employee data to visual insights.

## Features

- 100 employee records with 13 attributes
- 7 sheets covering all HR analytics areas
- 170 live Excel formulas (COUNTIFS, AVERAGEIFS, SUMIFS, MINIFS, MAXIFS, INDEX-MATCH)
- 4 embedded charts (Bar, Pie, Line, Horizontal Bar)
- Conditional formatting with color scales and data bars
- Dynamic employee lookup tool using INDEX-MATCH
- Attrition analysis by department and experience band
- Salary band analysis (min, max, avg, total payroll)

## Tech Stack

- Python 3.x
- openpyxl library
- Microsoft Excel

## Sheets

| Sheet | Description |
|---|---|
| Cover | Project overview and navigation |
| Employee_Data | Raw data — 100 employees, 13 fields |
| Summary | KPI cards + department breakdown |
| Charts | 4 embedded visual charts |
| Attrition_Analysis | Resignation rate by dept & experience |
| Employee_Lookup | Dynamic card using INDEX-MATCH |
| Salary_Analysis | Salary bands by department |

## Key Excel Skills Demonstrated

COUNTIFS · AVERAGEIFS · SUMIFS · MINIFS · MAXIFS · INDEX-MATCH · Conditional Formatting · Data Validation · Chart Creation · Dashboard Design

## How to Run

```bash
pip install openpyxl
python build_hr_dashboard.py
```

## Resume Bullet

> Built an HR Analytics Dashboard in Excel analyzing 100+ employee records across 7 departments using COUNTIFS, INDEX-MATCH, and Conditional Formatting to surface attrition, salary, and performance insights. Automated report generation using Python (openpyxl).
