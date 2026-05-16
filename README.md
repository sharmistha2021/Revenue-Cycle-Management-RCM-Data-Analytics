# 🏥 RCM Patient Billing Dashboard — Power BI

A two-page Revenue Cycle Management (RCM) dashboard built in Power BI using a synthetic dataset of 500 patient billing records. This is an early learning project — numbers are randomized and may not reflect real-world RCM benchmarks. Built to understand the domain, practice DAX and learn storytelling through data.

---

## 📊 Dashboard Pages

### Page 1 — The Revenue Story
> *How is our revenue performing?*

- KPI cards with MOM % change (Total Billed, Total Collected, Collection Rate, Avg Days Outstanding, Total Patients, Revenue per Patient)
- Monthly Billed vs Collected bar chart
- Claim Status donut chart
- Appeal Win Rate by Denial Reason
- Total Collected by Department
- Collection Rate by Payer with conditional formatting (green / amber / red)

### Page 2 — The Risk Story
> *Where is revenue at risk?*

- KPI strip (Denied Amount, Denial Rate, AR Over 90 Days, Total Outstanding)
- Denial Rate by Service Type
- Total Patients by Denial Reason
- AR Outstanding by Aging Bucket (stacked bar — green to red)
- Denial Rate trend by Month
- Denial Rate by Payer / Insurance

---

## 🗂️ Repository Structure

```
rcm-dashboard/
│
├── dataset/
│   └── rcm_patient_billing_ar.xlsx      # Synthetic dataset (500 rows)
│
├── dashboard/
│   └── RCM_Dashboard.pbix               # Power BI dashboard file
│
├── screenshots/
│   ├── page1_revenue_story.png          # Page 1 screenshot
│   └── page2_risk_story.png             # Page 2 screenshot
│
└── README.md
```

---

## 📁 Dataset

> ⚠️ **This is 100% synthetic data** generated in Python. No real patient information was used. Numbers are randomized and may not reflect real-world RCM benchmarks.

**500 patient records** across **2 sheets:**

| Column | Description |
|---|---|
| Patient ID | Unique identifier (PT10001–PT10500) |
| Patient Name | Randomly generated name |
| Admission Date | Oct 2025 – Mar 2026 |
| Discharge Date | Admission + random length of stay (0–14 days) |
| Service Type | Inpatient, Outpatient, Surgery, Emergency, etc. |
| Department | Cardiology, Oncology, Emergency, etc. |
| Payer / Insurance | Medicare, Medicaid, BCBS, Aetna, United Healthcare, Cigna, Self-Pay |
| Amount Billed ($) | $500 – $85,000 |
| Amount Paid ($) | % of billed based on claim status |
| Outstanding Balance ($) | Billed − Paid |
| Claim Status | Paid, Pending, Denied, Appealed, Partially Paid |
| Denial Reason | Prior Auth, Eligibility, Coding Error, Duplicate, Timely Filing, etc. |
| Payment Date | Discharge + 5–60 days (paid claims only) |
| Days Outstanding | Days from discharge to payment or today |
| AR Aging Bucket | 0–30d, 31–60d, 61–90d, 91–120d, 120+ days |
| Appeal Outcome | Won / Lost / N/A (appealed claims only) |

**Payer distribution** mirrors real-world hospital payer mix:
- Medicare 34% · Medicaid 21% · Commercial 28% · Self-Pay 11% · Other 6%

---

## 🔢 DAX Measures

### Base Measures
```dax
Total Billed = SUM('Patient_Billing_AR'[Amount Billed ($)])

Total Collected = SUM('Patient_Billing_AR'[Amount Paid ($)])

Total Outstanding = SUM('Patient_Billing_AR'[Outstanding Balance ($)])

Collection Rate = DIVIDE([Total Collected], [Total Billed], 0)

Denial Rate =
DIVIDE(
    COUNTROWS(FILTER('Patient_Billing_AR',
    'Patient_Billing_AR'[Claim Status] = "Denied")),
    COUNTROWS('Patient_Billing_AR'),
    0
)

Avg Days Outstanding = AVERAGE('Patient_Billing_AR'[Days Outstanding])
```

### MOM % Change (auto-detects latest month)
```dax
MOM Collected =
VAR _LastDate = MAX('Calendar'[Date])
VAR _CM = CALCULATE([Total Collected],
            DATESBETWEEN('Calendar'[Date],
                DATE(YEAR(_LastDate), MONTH(_LastDate), 1),
                _LastDate))
VAR _PM = CALCULATE([Total Collected],
            DATESBETWEEN('Calendar'[Date],
                DATE(YEAR(EDATE(_LastDate,-1)), MONTH(EDATE(_LastDate,-1)), 1),
                EDATE(_LastDate,-1)))
VAR _perc = DIVIDE(_CM - _PM, _PM)
VAR _percformat = FORMAT(_perc, "+0.0%;-0.0%;0.0%")
RETURN
IF(ISBLANK(_PM), BLANK(),
    IF(_perc > 0,
        UNICHAR(129157) & " " & _percformat,
        UNICHAR(129158) & " " & _percformat))
```

### Conditional Formatting Colors
```dax
CF Collected =
VAR _p = [MOM Collected %]
RETURN SWITCH(TRUE(),
    _p > 0,  "#1D9E75",
    _p < 0,  "#E24B4A",
    "#888780")

CF Denial =
VAR _p = [MOM Denial %]
RETURN SWITCH(TRUE(),
    _p < 0,  "#1D9E75",   -- lower denial = good
    _p > 0,  "#E24B4A",
    "#888780")
```

### AR Aging
```dax
AR Over 90 Days =
CALCULATE([Total Outstanding],
    'Patient_Billing_AR'[AR Aging Bucket] IN {"91-120 days", "120+ days"})

AR Over 90 Days % = DIVIDE([AR Over 90 Days], [Total Outstanding], 0)
```

### Appeal Win Rate
```dax
Appeal Win Rate =
DIVIDE(
    CALCULATE(COUNTROWS('Patient_Billing_AR'),
        'Patient_Billing_AR'[Appeal Outcome] = "Won"),
    CALCULATE(COUNTROWS('Patient_Billing_AR'),
        'Patient_Billing_AR'[Appeal Outcome] <> "N/A"),
    0
)
```

---

## 🗓️ Calendar Table

```dax
Calendar =
VAR StartDate = MIN('Patient_Billing_AR'[Admission Date])
VAR EndDate   = MAX('Patient_Billing_AR'[Discharge Date])
RETURN
ADDCOLUMNS(
    CALENDAR(StartDate, EndDate),
    "Year",         YEAR([Date]),
    "Month Number", MONTH([Date]),
    "Month Name",   FORMAT([Date], "MMMM"),
    "Month Short",  FORMAT([Date], "MMM"),
    "Quarter",      "Q" & QUARTER([Date]),
    "Week Number",  WEEKNUM([Date]),
    "Day Name",     FORMAT([Date], "DDDD"),
    "Year-Month",   FORMAT([Date], "YYYY-MM"),
    "MonthYear",    FORMAT([Date], "MMM YYYY")
)
```

**Relationships:**
- `Calendar[Date]` → `Patient_Billing_AR[Admission Date]` ✅ Active
- `Calendar[Date]` → `Patient_Billing_AR[Discharge Date]` ❌ Inactive
- `Calendar[Date]` → `Patient_Billing_AR[Payment Date]` ❌ Inactive

---

## 🛠️ Tools Used

| Tool | Purpose |
|---|---|
| Python + openpyxl | Generated synthetic dataset |
| Power BI Desktop | Dashboard design and DAX |
| DAX | All measures and calendar table |
| Excel | Dataset storage |

---

## ⚠️ Known Limitations

- Dataset is synthetic — real RCM data behaves differently
- Collection Rate includes pending claims in denominator — slightly understated
- Appeal Outcome column was generated using `RAND()` — win rates are approximations
- Billing Lag not included — requires Claim Submission Date column
- Month ordering on charts requires Calendar[Month Name] sorted by Month Number

---

## 📸 Screenshots

### Page 1 — The Revenue Story
![Revenue Story](screenshots/rcm_revenue.png)

### Page 2 — The Risk Story
![Risk Story](screenshots/rcm_risk.png)

---

## 🚀 How To Use

1. Download `rcm_patient_billing_ar.xlsx` from the `dataset/` folder
2. Open `RCM_Dashboard.pbix` in Power BI Desktop
3. If prompted, update the data source path to point to your local Excel file
4. Click **Refresh** — all visuals will update automatically

---

## 👤 Author

Built as a learning project to understand Revenue Cycle Management (RCM) in healthcare analytics.

Open to feedback — especially from anyone working in healthcare analytics or RCM. Still a lot to learn and improve!

---

## 📌 Disclaimer

> All data in this project is **100% synthetic** and was generated programmatically for learning purposes only. No real patient data was used. This project is not intended to represent actual clinical or financial performance of any healthcare organization.
