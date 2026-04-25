# 📊 eBay E-Commerce Sales Dashboard (Excel)

> **Raw Messy CSV → Executive-Ready Interactive Dashboard**  
> 13,000 rows | 5 Product Categories | 12 US Cities | Full Year 2024

---

## 🎯 Project Summary

A real-world Excel portfolio project that transforms a production-realistic messy eBay sales dataset into a fully interactive executive dashboard — using only **Microsoft Excel**. No Python. No SQL. No Power BI.

This project demonstrates the complete data analyst workflow:

```
Raw Messy CSV → Data Cleaning → PivotTable Analysis → Advanced Formulas → Interactive Dashboard
```

---

## 📁 Repository Structure

```
ebay-sales-dashboard/
│
├── 📄 README.md                         ← You are here
├── 📊 Dashboard.xlsx                    ← Main interactive Excel workbook ⭐
├── 📂 Raw-Data.csv                      ← Original messy dataset (before)
├── 📂 Clean-Data.xlsx                   ← Cleaned dataset (your work proof)
├── 📄 eBay_Business_Problem_Statement.pdf  ← Business context document
├── 📄 eBay_Project_Overview.pdf         ← Full project walkthrough PDF
└── 📂 Screenshots/
    └── dashboard-preview.png            ← Dashboard screenshot
```

---

## ❌ The Problem — What Was Wrong With the Raw Data

The raw CSV had **8 categories of data quality issues** across 13,000 rows:

| # | Problem | Example (Messy) | Example (Fixed) |
|---|---------|-----------------|-----------------|
| 1 | Mixed date formats | `1/15/25`, `01-15-2024`, `2024/01/15` | `2024-01-15` |
| 2 | Customer name + city combined | `"Alice - NYC"` | Name: `Alice` / City: `NYC` |
| 3 | Mixed price formats | `$34.99`, `35`, `34.9`, `N/A` | `34.99` |
| 4 | Inconsistent status casing | `delivered`, `Delivered`, `RETURNED` | `Valid` / `Invalid` |
| 5 | Mixed OrderID formats | `1001`, `ORD-1002` | `1001`, `1002` |
| 6 | Missing / NULL values | Blank Qty, N/A Price, empty Status | `0` via `IFERROR()` |
| 7 | Duplicate records | Same OrderID + Date appearing 2–3x | Removed via Remove Duplicates |
| 8 | Quoted product names | `'"T-Shirt"'` | `T-Shirt` |

---

## ✅ The Solution — Excel Cleaning Steps

Every issue was fixed using **native Excel tools only**:

| Step | Problem Solved | Excel Method | Formula Used |
|------|---------------|-------------|--------------|
| 1 | Split Customer + City | Text to Columns → Dash delimiter | `=TRIM(LEFT(A2,FIND("-",A2)-1))` |
| 2 | Standardize dates | DATEVALUE + Paste Values | `=DATEVALUE(TEXT(A2,"YYYY-MM-DD"))` |
| 3 | Remove $ from prices | Find & Replace → Ctrl+H | `=VALUE(SUBSTITUTE(G2,"$",""))` |
| 4 | Fix status casing | IF + LOWER | `=IF(LOWER(I2)="delivered","Valid","Invalid")` |
| 5 | Fix OrderID format | SUBSTITUTE + VALUE | `=VALUE(SUBSTITUTE(A2,"ORD-",""))` |
| 6 | Handle NULL values | IFERROR wrapper | `=IFERROR(VALUE(G2),0)` |
| 7 | Remove duplicates | Data → Remove Duplicates | OrderID + Date columns selected |
| 8 | Fix quoted products | Find & Replace | Remove extra quote characters |

### New Columns Created

| Column | Formula | Purpose |
|--------|---------|---------|
| `Revenue` | `=Qty * Price` | Core metric for all analysis |
| `Month` | `=TEXT(Date,"MMM-YY")` | Monthly grouping in PivotTables |
| `Status (clean)` | `=IF(LOWER(raw)="delivered","Valid","Invalid")` | Normalizes 5 variants → 2 values |
| `City` | Text to Columns split | Enables city-level slicer filtering |
| `Cust Total Revenue` | `=SUMIF($B:$B,B2,$J:$J)` | Lifetime revenue per customer |
| `Cust Tier` | `=INDEX(TierTable,MATCH(K2,Thresholds,1))` | Bronze / Silver / Gold classification |

---

## 🔄 PivotTables Built

**PivotTable 1 — Revenue by Category**

| Category | Orders | Revenue |
|----------|--------|---------|
| Electronics | 2,513 | $3,378,103 |
| Office | 2,616 | $633,753 |
| Clothing | 2,607 | $632,782 |
| Home | 2,594 | $620,761 |
| Sports | 2,670 | $479,493 |
| **Total** | **13,000** | **$5,744,894** |

**PivotTable 2 — Monthly Trends**  
12-month revenue and order count tracking (Jan–Dec 2024). Peak month: **March 2024 at $534,750**.

**PivotTable 3 — Top 10 Customers**

| Customer | Revenue | Tier |
|----------|---------|------|
| Charlie | $340,431 | 🥇 Gold |
| Ivy | $339,833 | 🥇 Gold |
| Tina | $322,714 | 🥇 Gold |
| Paul | $316,435 | 🥇 Gold |
| Eve | $308,798 | 🥈 Silver |

**PivotTable 4 — Product Matrix (Category × Month Heatmap)**  
Cross-tab of all 5 categories across 12 months with color-scale conditional formatting.

> All 4 PivotTables connected via **Slicers** (Category / Status / City)

---

## 🧮 Advanced Formulas

### Dynamic KPI Formulas

```excel
// Total Revenue — delivered orders only
=SUMIFS('eBay-Cleaned CSV'!J:J, 'eBay-Cleaned CSV'!I:I, "Valid")

// Orders Count
=SUMPRODUCT(('eBay-Cleaned CSV'!I:I="Valid")*1)

// Average Order Value
=TotalRevenue / OrdersCount

// Month-over-Month Growth
=(CurrentMonthRev - PrevMonthRev) / PrevMonthRev
```

### Customer Tier — INDEX-MATCH

```excel
// Step 1: Customer lifetime revenue
=SUMIF($B:$B, B2, $J:$J)

// Step 2: Assign tier using INDEX-MATCH
=INDEX(TierTable!$B$2:$B$4, MATCH(K2, TierTable!$A$2:$A$4, 1))
```

**Tier Thresholds** (based on actual data range $242K–$340K):

| Tier | Threshold | Revenue Range |
|------|-----------|---------------|
| 🥉 Bronze | 0 | $242,560 – $269,999 |
| 🥈 Silver | 270,000 | $270,000 – $309,999 |
| 🥇 Gold | 310,000 | $310,000 – $340,431 |

> Thresholds are **percentile-based** — not hardcoded — so they adapt to the actual data distribution.

---

## 📈 Dashboard KPI Results

| KPI | Value | Formula Method |
|-----|-------|----------------|
| 💰 Total Revenue | **$5,744,894** | SUMIFS on Valid orders |
| 📦 Total Orders | **13,000** | SUMPRODUCT count |
| 🛒 Avg Order Value | **$457.97** | Revenue ÷ Orders |
| 📈 MoM Growth | **+44.2%** | ($93K → $134K) |
| 🏆 Top Category | **Electronics** | 59% of total revenue |
| ⭐ Top Customer | **Charlie** | $340,431 lifetime revenue |

---

## 🖥️ Dashboard Features

```
┌─────────────────────────────────────────────────────────────┐
│         eBay Sales Dashboard — 2024                         │
├──────────┬──────────┬──────────┬──────────┬────────────────┤
│ Revenue  │  Orders  │   AOV    │  Growth  │  Top Category  │  ← KPI Cards
│$5,744,894│  13,000  │ $457.97  │  +44.2%  │  Electronics   │
├──────────┴──────────┴──────────┴──────────┴────────────────┤
│ Revenue Trend (Line+Bar) │ Category Donut │ Top 10 Customers│  ← Charts
├─────────────────────────────────────────────────────────────┤
│         Category × Month Heatmap (Color Scale)              │  ← Matrix
├─────────────────────────────────────────────────────────────┤
│  Slicers: Category │ Status │ City │ Customer Tier          │  ← Filters
└─────────────────────────────────────────────────────────────┘
```

**Dashboard Highlights:**
- ✅ 6 dynamic KPI cards with green/red conditional formatting
- ✅ Revenue Trend — Line + Bar combo chart (12 months)
- ✅ Category Donut chart with % data labels
- ✅ Top 10 Customers horizontal bar chart
- ✅ Category × Month heatmap with color scale
- ✅ 4 slicers connected to ALL PivotTables simultaneously
- ✅ Blue theme (#2E5A8D primary / #F4F7FA background)
- ✅ Freeze panes + 90% zoom for clean presentation

---

## 🗂️ Excel Workbook Sheets

| Sheet | Contents |
|-------|----------|
| `eBay-Cleaned CSV` | 13,000 cleaned rows — production-ready data |
| `Pivot tables` | All 4 PivotTables with slicers |
| `Kpi's` | Dynamic KPI formula cells |
| `Dashboard` | Final interactive executive dashboard |
| `Tiers` | Tier threshold lookup table |

---

## 💼 Skills Demonstrated

```
✅ Data Cleaning          — Text to Columns, Find & Replace, IFERROR, DATEVALUE
✅ Data Transformation    — SUMIF, SUBSTITUTE, VALUE, TEXT, TRIM, IF, LOWER
✅ PivotTable Analysis    — Revenue, trends, rankings, matrix cross-tab
✅ Advanced Formulas      — SUMIFS, SUMPRODUCT, INDEX-MATCH, array logic
✅ Dashboard Design       — KPI cards, charts, heatmap, conditional formatting
✅ Interactivity          — Slicers connected across all PivotTables
✅ Business Thinking      — Tier segmentation, MoM growth, AOV analysis
```

---

## 📊 Key Business Insights

1. **Electronics dominates** — 59% of total revenue ($3.37M of $5.74M) with highest AOV
2. **Top 3 customers (Charlie, Ivy, Tina)** — represent significant revenue concentration → VIP retention priority
3. **MoM growth +44.2%** — strong demand momentum ($93K → $134K)
4. **March 2024 peak** — highest month at $534,750 → optimal for inventory stocking
5. **Sports & Home underperform** — opportunity for cross-sell and bundle campaigns
6. **Gold tier customers** (revenue > $310K) → need dedicated retention programs
7. **12 cities tracked** — varying AOVs suggest geo-targeted marketing opportunities

---

## 🚀 How to Use This Dashboard

1. **Download** `Dashboard.xlsx`
2. **Enable editing** if prompted by Excel
3. **Click any slicer** (Category / Status / City / Tier) — watch all charts update live
4. **Hover over charts** for tooltips
5. **Check the KPI sheet** for formula breakdown

> 💡 Best viewed in **Microsoft Excel 2016 or later** at **90% zoom**

---

## 📬 Connect

Feel free to reach out or connect on LinkedIn if you'd like to discuss this project!

---

*Built with Microsoft Excel | eBay Sales Dataset | 2024*
