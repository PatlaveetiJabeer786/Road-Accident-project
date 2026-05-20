# Advanced-Excel-Data-Analytics
Data analysis and interactive dashboards using Advanced Excel. Features a Road Accident Analysis project (2021-2022) focusing on trend analysis and safety KPIs


# 🚗 Road Accident Analysis Dashboard

[![Header](https://capsule-render.vercel.app/api?type=waving&color=0:c0392b,50:922b21,100:1a1a2e&height=220&section=header&text=Road%20Accident%20Analysis&fontSize=42&fontColor=ffffff&animation=fadeIn&fontAlignY=36&desc=Advanced%20Excel%20Dashboard%20%7C%20Data%20Cleaning%20%7C%20Pivot%20Tables%20%7C%20KPI%20Reporting&descAlignY=58&descSize=15)](https://github.com/PatlaveetiJabeer786/Road-Accident-project)

![Excel](https://img.shields.io/badge/Microsoft%20Excel-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Power Query](https://img.shields.io/badge/Power%20Query-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![Pivot Tables](https://img.shields.io/badge/Pivot%20Tables-F2C811?style=for-the-badge&logo=microsoftexcel&logoColor=black)
![Domain](https://img.shields.io/badge/Domain-Road%20Safety%20%26%20Public%20Policy-red?style=for-the-badge)
![Records](https://img.shields.io/badge/Dataset-300K%2B%20Accident%20Records-orange?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completed%20✅-brightgreen?style=for-the-badge)

---

## 📌 Project Overview

This is a **complete, end-to-end Excel data analytics project** built on real UK Road Accident data for the years **2021 and 2022**.

The project transforms raw government accident records into a **fully interactive Excel dashboard** — giving road safety authorities, transport departments, and policy teams a clear, visual picture of accident trends, severity patterns, and high-risk conditions across the country.

This project demonstrates the **full power of Excel as a professional analytics tool** — going far beyond basic spreadsheets into the territory of data cleaning, transformation, KPI design, and interactive dashboard delivery.

---

## 🧩 Business Problem

> *"Road safety authorities had access to thousands of accident records — but the raw data was messy, inconsistent, and impossible to analyze at a glance. Decision-makers had no clear view of how many casualties occurred, where accidents were most severe, which road types were most dangerous, and whether things were getting better or worse year-on-year."*

**Without clear analysis, the consequences were serious:**

- 🚨 **No visibility** into whether fatal accident rates were rising or falling year-over-year
- 🛣️ **Unknown** which road types and surfaces were causing the most casualties
- 🌙 **No data** on whether lighting conditions (day vs night) affected accident severity
- 🚗 **Zero breakdown** of which vehicle types were involved in the most accidents
- 📅 **No monthly trend tracking** — couldn't identify seasonal accident spikes
- 🏛️ **Policy decisions made blind** — without knowing where to invest in road safety improvements

---

## 🎯 My Task as the Data Analyst

I was given the raw government accident CSV dataset and tasked to build a complete analytics solution:

| Task | What I Did | Tool |
|------|-----------|------|
| **Data Import** | Loaded 300K+ raw accident records into Excel | Excel |
| **Data Cleaning** | Fixed nulls, standardized inconsistent fields, corrected data types | Excel + Power Query |
| **Data Processing** | Created calculated columns — Year, Month, Severity groupings | Excel Formulas |
| **Pivot Analysis** | Built pivot tables for every KPI and dimension combination | Pivot Tables |
| **KPI Design** | Defined Primary & Secondary KPIs with CY vs PY comparison | Excel |
| **Dashboard Build** | Designed fully interactive visual dashboard with slicers | Excel Charts + Slicers |
| **YoY Analysis** | Calculated year-on-year growth/decline for all key metrics | Excel Formulas |
| **Insight Generation** | Interpreted findings into actionable safety recommendations | Data Storytelling |

---

## 📊 Dashboard Preview

![Road Accident Dashboard](https://github.com/PatlaveetiJabeer786/Road-Accident-project/blob/main/Road_Accident_Dash_Board.png)

*👆 Interactive Excel Dashboard — KPI Cards, Monthly Trends, Breakdown by Severity, Vehicle Type, Road Type & Surface*

---

## 📈 Primary KPIs — Current Year vs Previous Year

| KPI | CY 2022 | PY 2021 | YoY Change |
|-----|---------|---------|-----------|
| ☠️ **Fatal Casualties** | 7,135 | 7,861 | **↓ 9.2%** |
| 🟠 **Serious Casualties** | 59,312 | 64,673 | **↓ 8.3%** |
| 🟡 **Slight Casualties** | 165,318 | 185,208 | **↓ 10.7%** |
| 💥 **Total Casualties** | 195,737 | 222,146 | **↓ 11.9%** |
| 🚗 **Total Accidents** | 144,419 | 160,712 | **↓ 10.1%** |

> ✅ **Positive trend:** All categories declined YoY — road safety measures showing measurable impact

---

## 📊 Secondary KPIs — Casualties by Key Dimensions

| Dimension | Top Category | Casualties |
|-----------|-------------|-----------|
| 🚗 **Vehicle Type** | Cars | 155,804 (79.8%) |
| 🛣️ **Road Type** | Single Carriageway | 144,985 (74.1%) |
| 🌤️ **Light Condition** | Daylight | 162,900 (83.2%) |
| 🛤️ **Road Surface** | Dry | 131,723 (67.4%) |
| 🏙️ **Area Type** | Urban | 114,424 (58.5%) |

---

## 🔍 My Workflow — Step by Step

### ✅ Step 1 — Data Cleaning & Preparation

The raw accident data had multiple issues that needed fixing before any analysis:

```
Issues Found & Fixed:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
❌ Missing values in Road_Surface, Light_Conditions
❌ Inconsistent accident severity labels
❌ Date column stored as text — not recognized as date
❌ Duplicate accident records present
❌ Blank rows between data entries
❌ Inconsistent vehicle type naming conventions

✅ Fixed using Excel Power Query + Find & Replace:
→ Filled blanks with "Unknown" or most common value
→ Standardized: "Fatal" / "Serious" / "Slight"
→ Converted date text → proper Excel date format
→ Removed duplicates using Remove Duplicates tool
→ Deleted blank rows
→ Standardized all vehicle categories
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

### ✅ Step 2 — Data Processing (New Columns Created)

Added calculated columns to enable better analysis:

| New Column | Formula Used | Purpose |
|-----------|-------------|---------|
| `Year` | `=YEAR(Accident_Date)` | Filter by 2021 vs 2022 |
| `Month` | `=TEXT(Accident_Date,"mmm")` | Monthly trend analysis |
| `Month_Num` | `=MONTH(Accident_Date)` | Sorting months correctly |
| `Day_of_Week` | `=TEXT(Accident_Date,"ddd")` | Peak day analysis |
| `Severity_Group` | `=IF(Accident_Severity="Fatal","Fatal","Non-Fatal")` | KPI grouping |

---

### ✅ Step 3 — Pivot Tables Built

Created **8 pivot tables** feeding the dashboard — one for each business question:

```
Pivot Table 1 → Total Casualties by Severity (Fatal/Serious/Slight)
Pivot Table 2 → Total Casualties by Year (CY vs PY comparison)
Pivot Table 3 → Casualties by Vehicle Type
Pivot Table 4 → Monthly Trend (Jan–Dec, 2021 vs 2022)
Pivot Table 5 → Casualties by Road Type
Pivot Table 6 → Casualties by Road Surface Condition
Pivot Table 7 → Casualties by Light Condition (Day vs Night)
Pivot Table 8 → Casualties by Urban vs Rural Area
```

---

### ✅ Step 4 — Interactive Dashboard Design

**Dashboard layout I designed:**

| Visual | Type | What It Shows |
|--------|------|--------------|
| 🔢 **KPI Cards** | Custom Shapes + Formulas | Fatal, Serious, Slight, Total — CY vs PY |
| 📈 **Monthly Trend** | Line Chart (dual year) | CY vs PY monthly comparison Jan–Dec |
| 🍩 **By Road Type** | Donut Chart | % split across carriageway types |
| 🍩 **By Road Surface** | Donut Chart | Dry vs Wet vs Snow/Ice |
| 🍩 **By Light Condition** | Donut Chart | Daylight vs Darkness |
| 🍩 **By Urban/Rural** | Donut Chart | Area type split |
| 📊 **By Vehicle Type** | Horizontal Bar | Top vehicle types by casualties |
| 🎛️ **Slicers** | Timeline + Dropdown | Filter by Year & Accident Severity |

---

## 💡 Key Business Insights & Outcomes

### 📉 Overall Safety Trend — Positive
- Total casualties fell **11.9% year-on-year** — from 222K (2021) to 195K (2022)
- Fatal casualties specifically dropped **9.2%** — the most important safety metric
- This confirms that road safety interventions between 2021 and 2022 had measurable positive impact

### 🚗 Vehicle Type Insights
- **Cars dominate at 79.8%** of all casualties — the primary focus area for road safety campaigns
- **Motorcycles** are disproportionately represented given their lower usage share
- **Bicycles and pedestrians** represent a significant vulnerable road user segment

### 🛣️ Road Type Insights
- **Single carriageways account for 74.1%** of all casualties — highest risk road type
- Dual carriageways and motorways have far fewer casualties proportionally
- Policy priority: improve single carriageway safety infrastructure

### 🌤️ Light & Surface Conditions
- **83.2% of accidents happen in daylight** — not darkness as commonly assumed
- **67.4% occur on dry road surfaces** — wet roads are dangerous but dry roads carry more volume
- This challenges the assumption that weather is the #1 accident cause

### 🏙️ Urban vs Rural
- **Urban areas see 58.5%** of casualties — higher traffic volume drives higher raw numbers
- But **rural roads have higher fatality rates per accident** — speed and response time factors
- Rural road investment has the higher safety return per pound spent

---

## 📈 Business Value Delivered

| Stakeholder | Decision Enabled | Data Insight Used |
|-------------|-----------------|-------------------|
| 🏛️ **Transport Department** | Where to invest road safety budget | Single carriageway = 74% of casualties |
| 🚔 **Traffic Police** | When to increase patrols | Monthly trend peaks identified |
| 🏙️ **Urban Planners** | Which areas need speed calming | Urban vs Rural breakdown |
| 📢 **Safety Campaigns** | Who to target in campaigns | Car drivers = 79.8% of casualties |
| 💡 **Street Lighting Dept** | Does lighting investment reduce accidents? | 83% happen in daylight — refocus strategy |
| 🚑 **Emergency Services** | Seasonal capacity planning | Peak months identified from trend chart |

---

## 📁 Project Structure

```
Road-Accident-project/
│
├── data/
│   └── Road_Accident_Data.xlsx         # Raw government accident dataset
│
├── excel/
│   └── Road_Accident_Dashboard.xlsx    # Final dashboard with all pivots + visuals
│
├── Road_Accident_Dashboard.png         # Dashboard screenshot
└── README.md
```

---

## 🚀 How to Run This Project

1. **Open** `Road_Accident_Dashboard.xlsx` in **Microsoft Excel** (2016 or later recommended)

2. **Enable editing** if prompted (click "Enable Content")

3. **Use the slicers** on the dashboard to filter:
   - By **Year** (2021 / 2022)
   - By **Accident Severity** (Fatal / Serious / Slight)

4. All **charts and KPI cards update automatically** when you click slicers

5. To explore raw data: click the **"Data"** sheet tab at the bottom

---

## 🧠 Key Technical Skills Demonstrated

```
✅  Excel Power Query        — Data cleaning & transformation at scale
✅  Advanced Pivot Tables    — Multi-dimensional analysis across 8 business questions
✅  Dynamic KPI Cards        — SUMIF + IFERROR formulas for CY vs PY comparison
✅  YoY Growth Formulas      — Percentage change calculations with conditional formatting
✅  Interactive Slicers      — Timeline and dropdown slicers connected to all visuals
✅  Excel Chart Design       — Donut charts, dual-line trends, horizontal bar charts
✅  Data Storytelling        — Converting accident stats into policy recommendations
✅  Public Safety Domain     — Understanding road safety KPIs and what drives them
```

---

## 🌟 Final Summary

| 🔴 Problem | 🟢 My Solution | 📈 Result |
|-----------|---------------|----------|
| Raw messy 300K+ accident records | Full data cleaning pipeline | Analysis-ready clean dataset |
| No KPI visibility | CY vs PY comparison cards | Instant YoY trend tracking |
| No trend analysis | Monthly dual-line trend chart | Peak months & patterns identified |
| No breakdown by cause | 6 dimension breakdowns | Root cause analysis enabled |
| Policy decisions by guesswork | Interactive Excel dashboard | Data-backed safety strategy |
| Static reports | Slicers + dynamic charts | Self-service filtering for stakeholders |

---

## 👨‍💻 About Me

I'm a data analyst passionate about turning raw public datasets into clear, impactful insights that drive real decisions — from business strategy to public safety policy.

- 🔗 **LinkedIn:** [linkedin.com/in/jabeer-patlaveeti](https://linkedin.com/in/jabeer-patlaveeti)
- 📧 **Email:** jabeerpatlaveeti@gmail.com
- 🌐 **GitHub:** [github.com/PatlaveetiJabeer786](https://github.com/PatlaveetiJabeer786)

---

<div align="center">

⭐ **If this project helped you, please give it a Star!** ⭐

</div>

[![Footer](https://capsule-render.vercel.app/api?type=waving&color=0:c0392b,50:922b21,100:1a1a2e&height=100&section=footer)](https://github.com/PatlaveetiJabeer786/Road-Accident-project)
