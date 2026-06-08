# 📊 دليل لوحة التحكم المالية CFO - Power BI
# 📊 CFO Financial Dashboard - Power BI Guide

---

## 🎯 نظرة عامة | Overview

هذا الدليل يشرح كيفية **ربط بيانات مالية خام مع Power BI** لإنشاء **لوحة تحكم CFO احترافية** تعرض:

This guide explains how to **connect raw financial data to Power BI** to create a **professional CFO dashboard** that displays:

✅ التحليلات المالية الشاملة | Comprehensive financial analysis
✅ المؤشرات الرئيسية KPI | Key Performance Indicators
✅ التنبؤات والسيناريوهات | Forecasts & Scenario Analysis
✅ تحليل النسب والمخاطر | Ratio & Risk Analysis

---

## 📁 هيكل الملفات | File Structure

```
financial-analysis/
├── Finance Data & Analysis/
│   ├── DASHBOARD.pbix                    ← ملف Power BI الرئيسي
│   ├── findata.xlsx                      ← البيانات المحللة
│   ├── findata Before analysis.xlsx      ← البيانات الخام الأصلية
│   └── POWER-BI-DASHBOARD-GUIDE.md      ← هذا الملف (التوثيق)
├── SKILL.md                             ← مواصفات خبير التحليل المالي
└── About Financial-Analysis Skill.md
```

---

## 🔗 الخطوة 1: تجهيز البيانات | STEP 1: Data Preparation

### 1.1 بيانات المصدر | Source Data

**الملف:** `findata.xlsx`

**الأوراق المطلوبة:**
- [ ] **Balance Sheet** - الميزانية العمومية
- [ ] **Income Statement** - قائمة الدخل
- [ ] **Cash Flow** - قائمة التدفق النقدي
- [ ] **Budget Data** - بيانات الميزانية (اختياري)
- [ ] **Historical Data** - البيانات التاريخية (3-5 سنوات)

### 1.2 متطلبات تنسيق البيانات | Data Format Requirements

```
الأعمدة الأساسية / Required Columns:

✓ Account Code          - رمز الحساب
✓ Account Name          - اسم الحساب (العربية والإنجليزية)
✓ Period / Month        - الفترة (صيغة: YYYY-MM)
✓ Amount                - المبلغ (بدون فاصل عملات)
✓ Currency              - العملة (USD, SAR, etc.)
✓ Statement Type        - نوع القائمة (BS, IS, CF)
✓ Category              - الفئة الرئيسية
✓ Subcategory           - الفئة الفرعية
```

### 1.3 تنظيف البيانات | Data Cleaning

```excel
✅ إزالة الصفوف الفارغة
✅ توحيد أسماء الأعمدة (بدون مسافات إضافية)
✅ تحويل التواريخ إلى صيغة موحدة (YYYY-MM-DD)
✅ إزالة القيم المكررة
✅ التحقق من توازن الميزانية العمومية
✅ التحقق من التدفق النقدي الختامي = نقد الميزانية
```

---

## 📥 الخطوة 2: ربط البيانات مع Power BI | STEP 2: Connect Data to Power BI

### 2.1 فتح Power BI Desktop

1. افتح **Power BI Desktop**
2. اختر **Get Data** → **Excel**
3. اختر ملف `findata.xlsx`
4. حدد جميع الأوراق (Sheets)
5. اضغط **Load & Transform**

### 2.2 تحويل البيانات | Data Transformation

في **Power Query Editor**:

```
1. إضافة عمود التاريخ (إذا لم يكن موجوداً):
   = Date.FromText([Period])

2. إضافة عمود السنة والشهر:
   Year = YEAR([Date])
   Month = MONTH([Date])
   MonthName = FORMAT([Date], "MMMM", "ar-SA")

3. إضافة عمود النوع (Actual vs Budget):
   Type = "Actual" (أو من عمود آخر)

4. التحقق من جودة البيانات:
   ✓ No null values في الأعمدة الأساسية
   ✓ Data types صحيحة (Number, Date, Text)
```

### 2.3 إنشاء جداول القاموس | Create Dimension Tables

**جدول الحسابات (Account Dimension):**
```
Account ID | Account Name | Category | Subcategory | Statement Type
```

**جدول الفترات (Time Dimension):**
```
Date | Year | Quarter | Month | WeekNum | Date_Key
```

---

## 🧮 الخطوة 3: إنشاء المقاييس DAX | STEP 3: Create DAX Measures

### 3.1 المقاييس الأساسية | Basic Measures

```dax
-- الإيرادات / Revenue
Revenue = SUM(Financials[Amount])
WHERE Financials[Account Code] = "4000"

-- تكلفة البضائع المباعة / COGS
COGS = SUM(Financials[Amount])
WHERE Financials[Account Code] IN ("5000", "5100")

-- مجمل الربح / Gross Profit
Gross_Profit = [Revenue] - [COGS]

-- مصاريف التشغيل / Operating Expenses
OpEx = SUM(Financials[Amount])
WHERE Financials[Category] = "Operating Expenses"

-- الدخل التشغيلي / EBIT
EBIT = [Gross_Profit] - [OpEx]

-- صافي الدخل / Net Income
Net_Income = SUM(Financials[Amount])
WHERE Financials[Account Code] = "8000"
```

### 3.2 مقاييس الهوامش | Margin Measures

```dax
-- هامش مجمل الربح / Gross Margin %
Gross_Margin_Pct = DIVIDE([Gross_Profit], [Revenue], 0) * 100

-- هامش التشغيل / Operating Margin %
Op_Margin_Pct = DIVIDE([EBIT], [Revenue], 0) * 100

-- هامش صافي الربح / Net Margin %
Net_Margin_Pct = DIVIDE([Net_Income], [Revenue], 0) * 100

-- EBITDA (الاستهلاك والإطفاء)
DA = SUM(Financials[Amount])
WHERE Financials[Account Code] IN ("6200", "6300")

EBITDA = [EBIT] + [DA]

-- هامش EBITDA / EBITDA Margin %
EBITDA_Margin_Pct = DIVIDE([EBITDA], [Revenue], 0) * 100
```

### 3.3 مقاييس النسب المالية | Financial Ratios

```dax
-- نسبة التداول / Current Ratio
Current_Assets = SUM(Financials[Amount])
WHERE Financials[Category] = "Current Assets"

Current_Liabilities = SUM(Financials[Amount])
WHERE Financials[Category] = "Current Liabilities"

Current_Ratio = DIVIDE([Current_Assets], [Current_Liabilities], 0)

-- نسبة السيولة السريعة / Quick Ratio
Quick_Assets = [Current_Assets] - SUM(Financials[Amount])
WHERE Financials[Account Code] = "1300" -- Inventory

Quick_Ratio = DIVIDE([Quick_Assets], [Current_Liabilities], 0)

-- نسبة الدين للملكية / Debt to Equity
Total_Debt = SUM(Financials[Amount])
WHERE Financials[Category] IN ("Current Liabilities", "Long-Term Debt")

Total_Equity = SUM(Financials[Amount])
WHERE Financials[Category] = "Equity"

Debt_to_Equity = DIVIDE([Total_Debt], [Total_Equity], 0)

-- العائد على الملكية / ROE
ROE = DIVIDE([Net_Income], [Total_Equity], 0) * 100

-- العائد على الأصول / ROA
Total_Assets = SUM(Financials[Amount])
WHERE Financials[Category] IN ("Current Assets", "Non-Current Assets")

ROA = DIVIDE([Net_Income], [Total_Assets], 0) * 100
```

### 3.4 مقاييس المقارنة | Comparison Measures

```dax
-- مقابل الفترة السابقة / Prior Period
Prior_Period_Revenue = 
CALCULATE(
    [Revenue],
    DATEADD(Dates[Date], -1, MONTH)
)

-- نمو سنوي / YoY Growth
YoY_Growth = 
DIVIDE(
    [Revenue] - CALCULATE([Revenue], DATEADD(Dates[Date], -12, MONTH)),
    CALCULATE([Revenue], DATEADD(Dates[Date], -12, MONTH)),
    0
) * 100

-- نمو ربعي / QoQ Growth
QoQ_Growth = 
DIVIDE(
    [Revenue] - CALCULATE([Revenue], DATEADD(Dates[Date], -3, MONTH)),
    CALCULATE([Revenue], DATEADD(Dates[Date], -3, MONTH)),
    0
) * 100

-- الفرق عن الميزانية / Variance vs Budget
Budget_Revenue = 
CALCULATE(
    [Revenue],
    Financials[Type] = "Budget"
)

Variance_Amount = [Revenue] - [Budget_Revenue]

Variance_Pct = DIVIDE([Variance_Amount], [Budget_Revenue], 0) * 100
```

---

## 🎨 الخطوة 4: تصميم صفحات Dashboard | STEP 4: Design Dashboard Pages

### 📄 الصفحة 1: النظرة العامة التنفيذية | Executive Overview

**التخطيط:**
```
┌─────────────────────────────────────────────────────────┐
│              FINANCIAL DASHBOARD - OVERVIEW             │
├─────────────────────────────────────────────────────────┤
│  KPI 1: Revenue  │  KPI 2: Gross Profit │ KPI 3: EBITDA │
│  KPI 4: Net Inc  │  KPI 5: FCF          │ KPI 6: Cash   │
├��────────────────────────────────────────────────────────┤
│  Revenue Trend (Line Chart) │ Margin Trend (Multi-line) │
├─────────────────────────────────────────────────────────┤
│  YoY Growth (Column)        │ Key Metrics KPI Table     │
└─────────────────────────────────────────────────────────┘
```

**العناصر:**
1. **KPI Cards** (6 بطاقات):
   - 💰 الإيرادات / Revenue
   - 📊 مجمل الربح / Gross Profit
   - 📈 EBITDA
   - 💵 صافي الدخل / Net Income
   - 💧 التدفق النقدي الحر / FCF
   - 💳 الرصيد النقدي / Cash Balance

2. **Line Chart** - اتجاه الإيرادات:
   ```
   X-axis: Months / Quarters
   Y-axis: Revenue Amount
   Legend: Current Year, Prior Year
   ```

3. **Multi-line Chart** - اتجاه الهوامش:
   ```
   Lines: Gross Margin %, EBITDA Margin %, Net Margin %
   Target: تصفية حسب المدة المختارة
   ```

4. **Matrix/Table** - جدول المؤشرات الرئيسية:
   ```
   Rows: Metric Name
   Columns: Current Period, Prior Period, YoY Change, Status
   ```

---

### 📄 الصفحة 2: تعمق في قائمة الدخل | P&L Deep Dive

**التخطيط:**
```
┌──────────────────────────────────────────────────────┐
│          INCOME STATEMENT ANALYSIS                  │
├──────────────────────────────────────────────────────┤
│  Matrix: P&L Statement (Full Detail)               │
│  Rows: Revenue, COGS, Gross Profit, OpEx, EBIT...  │
│  Cols: Current, Prior, Change %, Budget, Variance  │
├──────────────────────────────────────────────────────┤
│  Waterfall Chart: Revenue → COGS → GP → OpEx → EBIT│
├──────────────────────────────────────────────────────┤
│  Variance Chart    │    Top 10 Expenses (Donut)     │
└──────────────────────────────────────────────────────┘
```

**العناصر:**
1. **Income Statement Matrix**:
   ```
   Line Items: Revenue, COGS, Gross Profit, SG&A, D&A, EBIT, 
               Interest, Taxes, Net Income
   Columns: Yr-2, Yr-1, Current, YoY Growth %, Budget, Variance
   Formatting: Conditional (Red/Yellow/Green)
   ```

2. **Waterfall Chart** - جسر الربح:
   ```
   Flow: Revenue → COGS → Gross Profit → OpEx → EBIT
   Shows: Impact of each component on final result
   ```

3. **Variance Chart** - الانحرافات:
   ```
   Type: Diverging Bar
   Data: Actual vs Budget, Favorable/Unfavorable
   Color: Green (Favorable), Red (Unfavorable)
   ```

4. **Expense Breakdown** - توزيع المصاريف:
   ```
   Type: Donut Chart
   Top 10 Expense Categories with percentages
   ```

---

### 📄 الصفحة 3: الميزانية العمومية والنقد | Balance Sheet & Cash

**التخطيط:**
```
┌──────────────────────────────────────────────────────┐
│         BALANCE SHEET & CASH ANALYSIS               │
├──────────────────────────────────────────────────────┤
│  BS Composition    │    Asset Breakdown (100% Stack) │
│  (Stacked Bar)     │    Liability/Equity (100%)     │
├──────────────────────────────────────────────────────┤
│  Working Capital Trend     │    Cash Balance Gauge   │
│  (Line Chart)              │    (Card + Trend)       │
├──────────────────────────────────────────────────────┤
│  FCF Bridge (Waterfall)                            │
│  EBITDA → ΔWC → CapEx → Debt Service → FCF         │
└──────────────────────────────────────────────────────┘
```

**العناصر:**
1. **Balance Sheet Composition** - تكوين الميزانية:
   ```
   Type: 100% Stacked Bar
   Categories: Current Assets, Fixed Assets, Current Liab, LT Debt, Equity
   ```

2. **Working Capital Trend**:
   ```
   Type: Line Chart with Area
   Current Assets - Current Liabilities over time
   Target Reference Line
   ```

3. **Cash & Equivalents**:
   ```
   Type: Card + Trend Indicator
   Shows: Current Balance, Prior Period, Change %
   ```

4. **FCF Bridge**:
   ```
   Type: Waterfall
   Flow: EBITDA → Less: Tax → Plus: ΔWC → Less: CapEx → FCF
   ```

---

### 📄 الصفحة 4: النسب والمعايير | Ratios & Benchmarks

**التخطيط:**
```
┌──────────────────────────────────────────────────────┐
│           RATIO ANALYSIS DASHBOARD                  │
├──────────────────────────────────────────────────────┤
│  Liquidity Ratios Trend  │  Profitability Trend     │
│  (Line Chart)            │  (Line Chart)            │
├──────────────────────────────────────────────────────┤
│  Efficiency Ratios Trend │  Solvency Ratios Trend   │
│  (Line Chart)            │  (Line Chart)            │
├──────────────────────────────────────────────────────┤
│  Ratio Scorecard (Table with RAG)                  │
│  Metric | Current | Benchmark | Status             │
└──────────────────────────────────────────────────────┘
```

**العناصر:**

1. **Liquidity Ratios Trend**:
   ```
   Metrics: Current Ratio, Quick Ratio, Cash Ratio
   Reference Lines: Healthy Range Benchmarks
   Color: Green (Healthy), Amber (Caution), Red (Risk)
   ```

2. **Profitability Ratios Trend**:
   ```
   Metrics: Gross Margin %, Op Margin %, Net Margin %, ROE %, ROA %
   Benchmark: Industry Average (if available)
   ```

3. **Efficiency Ratios Trend**:
   ```
   Metrics: Asset Turnover, Inventory Turnover, DSO, DPO, CCC
   ```

4. **Solvency Ratios Trend**:
   ```
   Metrics: Debt-to-Equity, Debt Ratio, Interest Coverage, Altman Z-Score
   Safe Zones: Visual reference lines
   ```

5. **Ratio Scorecard** - لوحة الحالة:
   ```
   Table Format:
   Ratio Name | Current Value | Benchmark | Status (✅/⚠️/🔴)
   
   Conditional Formatting:
   - Green: Above benchmark or in healthy range
   - Yellow: Near threshold
   - Red: Below threshold or risky
   ```

---

### 📄 الصفحة 5: التوقعات والسيناريوهات | Forecast & Scenarios

**التخطيط:**
```
┌──────────────────────────────────────────────────────┐
│       FORECAST & SCENARIO ANALYSIS                  │
├──────────────────────────────────────────────────────┤
│  Revenue Forecast        │  EBITDA Forecast         │
│  (Line with Range)       │  (Line with Range)       │
├──────────────────────────────────────────────────────┤
│  Scenario Comparison (Grouped Bar)                 │
│  Best Case | Base Case | Worst Case                │
├──────────────────────────────────────────────────────┤
│  FCF Forecast    │    Risk Heat Map (Table)        │
└──────────────────────────────────────────────────────┘
```

**العناصر:**

1. **Revenue Forecast**:
   ```
   Type: Line Chart with Confidence Band
   Historical: Solid Line
   Forecast: Dashed Line + Shaded Range (80% CI)
   X-axis: Next 12 months / quarters
   ```

2. **Scenario Comparison**:
   ```
   Type: Grouped Bar Chart
   Categories: Revenue, EBITDA, Net Income, FCF
   Series: Best Case, Base Case, Worst Case
   ```

3. **FCF Forecast**:
   ```
   Type: Bar Chart
   Positive/Negative Colors
   Trend Line
   ```

4. **Risk Heat Map**:
   ```
   Table Format:
   Risk | Evidence (Metric) | Level (🔴/🟡/🟢) | Mitigation
   
   Example:
   Liquidity Risk | Current Ratio = 0.78 | 🔴 HIGH | Improve WC
   ```

---

## 🎚️ الخطوة 5: إضافة المرشحات والتفاعلات | STEP 5: Add Filters & Interactions

### 5.1 مرشحات عامة | Global Filters

```
┌─ Period Filter (Time Slicer)
│  └─ Monthly / Quarterly / Yearly selection
│
├─ Year Filter
│  └─ Multi-select: 2021, 2022, 2023, 2024, 2025
│
├─ Department Filter (اختياري)
│  └─ Multi-select for business units
│
└─ Currency Filter (إذا متعدد)
   └─ USD, SAR, EUR, etc.
```

### 5.2 التفاعلات بين المرئيات | Visual Interactions

```
Revenue Chart → filters other visuals
Clicking on a month → updates all dependent charts
Department Filter → cascades to all pages

Cross-Filtering:
✓ Revenue Chart ← → P&L Matrix
✓ Period Slicer → All Pages
✓ Year Filter → All Visuals
```

---

## 📋 الخطوة 6: التنسيق والعلامات التجارية | STEP 6: Formatting & Branding

### 6.1 ألوان موحدة | Color Scheme

```
الإيجابي / Positive    : 🟢 #06B91D (Green)
محايد / Neutral        : 🔵 #4472C4 (Blue)
تحذير / Warning        : 🟡 #FFC000 (Yellow)
سلبي / Negative        : 🔴 #E81B23 (Red)

الخط الأساسي / Primary Font: Segoe UI (Bold for headers)
الخط الثانوي / Secondary: Calibri (Regular)
```

### 6.2 تنسيق الأرقام | Number Formatting

```
العملات / Currencies     : $#,##0 or $#,##0.0M
النسب المئوية / Percentages: 0.0%
الكميات / Quantities     : #,##0
الفترات الزمنية / Periods: MMM-YY (Jan-24)
```

### 6.3 رأس الصفحة والتذييل | Header & Footer

```
┌─────────────────────────────────────────────────────┐
│ 🏢 COMPANY NAME - FINANCIAL DASHBOARD               │
│ Last Updated: [Dynamic Date]                        │
│ Report Period: [Selected Period]                    │
└─────────────────────────────────────────────────────┘

Footer:
Data Source: findata.xlsx | Confidential
```

---

## 🔄 الخطوة 7: التحديثات والصيانة | STEP 7: Updates & Maintenance

### 7.1 جدول التحديثات | Update Schedule

```
يومي / Daily         : Cash Balance, Latest Transactions
أسبوعي / Weekly      : Revenue, Top Expenses
شهري / Monthly       : Full P&L, Balance Sheet, Ratios
ربعي / Quarterly      : Forecasts, Scenario Analysis
سنوي / Annually      : Strategic Review, Benchmarking
```

### 7.2 عملية التحديث | Update Process

```
1. فتح ملف البيانات الخام
   └─ findata Before analysis.xlsx

2. إضافة البيانات الجديدة للفترة الحالية
   ├─ Income Statement data
   ├─ Balance Sheet data
   └─ Cash Flow data

3. تشغيل تحليل مالي
   └─ (باستخدام SKILL.md workflow)

4. حفظ البيانات المحللة
   └─ findata.xlsx

5. تحديث الاتصالات في Power BI
   └─ Refresh Data → Publish

6. التحقق من Dashboard
   └─ ✓ جميع المقاييس محدثة
   └─ ✓ لا توجد أخطاء في الحسابات
   └─ ✓ التنسيق صحيح
```

---

## 🆘 استكشاف الأخطاء | Troubleshooting

### المشكلة: البيانات لم تُحدّث | Data Not Refreshing

```
الحل:
1. Power BI Desktop: Home → Refresh
2. Power BI Service: Dataset → Refresh now
3. تحقق من اتصال البيانات (Get Data → Edit)
4. تأكد من أن ملف Excel لم يُفتح في جهاز آخر
```

### المشكلة: الأرقام غير صحيحة | Incorrect Numbers

```
الحل:
1. تحقق من الصيغ في جدول البيانات الأصلي
2. تأكد من توازن الميزانية العمومية
3. تحقق من صيغ DAX (F2 key)
4. استخدم Column Diagnostics للبحث عن الأخطاء
```

### المشكلة: الأداء بطيء | Slow Performance

```
الحل:
1. قلل عدد الصفوف المحملة (استخدم Filters)
2. قلل التفاصيل في المرئيات (Summarize)
3. حذف الأعمدة غير المستخدمة
4. استخدم Aggregations (Power BI Pro)
5. قسم البيانات إلى عدة ملفات (Sharding)
```

---

## 📚 الموارد الإضافية | Additional Resources

### ملفات مرجعية | Reference Files

```
📄 SKILL.md
   └─ مواصفات كاملة لخبير التحليل المالي
   └─ أطر عمل التحليل
   └─ معايير الكتابة والجودة

📄 About Financial-Analysis Skill.md
   └─ نظرة عامة على المهارة
   └─ الحالات الاستخدام
   └─ أمثلة عملية
```

### روابط مفيدة | Useful Links

- [Power BI Documentation](https://docs.microsoft.com/en-us/power-bi/)
- [DAX Function Reference](https://dax.guide/)
- [Financial Ratio Formulas](https://en.wikipedia.org/wiki/Financial_ratio)
- [Cash Flow Analysis](https://www.investopedia.com/terms/c/cashflow.asp)

---

## ✅ قائمة التحقق النهائية | Final Checklist

```
Dashboard Setup:
☑ جميع الأوراق متصلة بـ Power BI
☑ مقاييس DAX تم إنشاؤها وتجربتها
☑ الصفحات الخمس تعرض البيانات صحيحة
☑ المرشحات تعمل بشكل صحيح
☑ التنسيق متسق وموحد

Data Quality:
☑ لا توجد قيم فارغة في الأعمدة الأساسية
☑ الأرقام صحيحة ومعقولة
☑ التواريخ بصيغة موحدة
☑ العملات متسقة

Performance:
☑ وقت التحميل < 5 ثوان
☑ الاستجابة للمرشحات سريعة
☑ لا توجد رسائل خطأ

Documentation:
☑ جميع المقاييس موثقة
☑ الفتراضات واضحة
☑ التعليقات موجودة في الأعمدة المعقدة
```

---

## 📞 الدعم والمساعدة | Support

للمساعدة في:
- ✉️ مشاكل الاتصال بالبيانات
- ✉️ صيغ DAX معقدة
- ✉️ تصميم المرئيات
- ✉️ التنبؤات والسيناريوهات

**تواصل عبر:** GitHub Issues أو البريد الإلكتروني

---

**إعداد:** Salih Al-Sharafi | **التاريخ:** 2025-06-08 | **الإصدار:** 1.0

---

