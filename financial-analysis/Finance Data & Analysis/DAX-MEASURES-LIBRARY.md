# 🧮 مكتبة مقاييس DAX للـ Dashboard المالي CFO
# 📊 DAX Measures Library for CFO Financial Dashboard

---

## 📌 مقدمة | Introduction

هذا الملف يحتوي على **جميع مقاييس DAX** اللازمة لـ Power BI Dashboard المالي الشامل.

This file contains **all required DAX measures** for the comprehensive Power BI CFO Dashboard.

**كيفية الاستخدام:** | **How to Use:**
1. افتح Power BI Desktop | Open Power BI Desktop
2. انتقل إلى **Data** view | Go to **Data** view
3. اختر جدول البيانات الرئيسي | Select main data table
4. انقر بزر اليمين → **New Measure** | Right-click → **New Measure**
5. انسخ والصق الصيغة من هنا | Copy & paste the formula below

---

## 📊 الفئة 1: مقاييس قائمة الدخل | SECTION 1: Income Statement Measures

### 1.1 الإيرادات والمبيعات | Revenue & Sales

```dax
-- الإيرادات الإجمالية / Total Revenue
Revenue = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "4000",
    Financials[Statement Type], "IS"
)

-- الإيرادات - الفترة السابقة / Revenue - Prior Month
Prior_Revenue = 
CALCULATE(
    [Revenue],
    DATEADD(Dates[Date], -1, MONTH)
)

-- الإيرادات - السنة الماضية / Revenue - Prior Year
PY_Revenue = 
CALCULATE(
    [Revenue],
    DATEADD(Dates[Date], -12, MONTH)
)

-- الإيرادات الأخرى / Other Revenue
Other_Revenue = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "4100",
    Financials[Statement Type], "IS"
)
```

### 1.2 تكلفة البضائع المباعة | Cost of Goods Sold

```dax
-- تكلفة البضائع المباعة / COGS
COGS = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], {"5000", "5100", "5200"},
    Financials[Statement Type], "IS"
)

-- COGS - الفترة السابقة / COGS - Prior Period
Prior_COGS = 
CALCULATE(
    [COGS],
    DATEADD(Dates[Date], -1, MONTH)
)

-- COGS - السنة الماضية / COGS - Prior Year
PY_COGS = 
CALCULATE(
    [COGS],
    DATEADD(Dates[Date], -12, MONTH)
)
```

### 1.3 مجمل الربح والهامش | Gross Profit & Margin

```dax
-- مجمل الربح / Gross Profit
Gross_Profit = [Revenue] - [COGS]

-- مجمل الربح - السابق / Prior Gross Profit
Prior_Gross_Profit = [Prior_Revenue] - [Prior_COGS]

-- مجمل الربح - السنة الماضية / Prior Year Gross Profit
PY_Gross_Profit = 
CALCULATE(
    [Gross_Profit],
    DATEADD(Dates[Date], -12, MONTH)
)

-- هامش مجمل الربح / Gross Margin %
Gross_Margin_Pct = 
IFERROR(
    DIVIDE([Gross_Profit], [Revenue]) * 100,
    BLANK()
)

-- هامش مجمل الربح - السنة الماضية / Prior Year Gross Margin %
PY_Gross_Margin_Pct = 
CALCULATE(
    [Gross_Margin_Pct],
    DATEADD(Dates[Date], -12, MONTH)
)
```

### 1.4 مصاريف التشغيل | Operating Expenses

```dax
-- إجمالي مصاريف التشغيل / Total Operating Expenses
OpEx = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Operating Expenses",
    Financials[Statement Type], "IS"
)

-- مصاريف البيع والعمومية والإدارية / Selling, General & Administrative
SGA = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], {"6100", "6200"},
    Financials[Statement Type], "IS"
)

-- نسبة SGA من الإيرادات / SG&A as % of Revenue
SGA_to_Revenue = 
IFERROR(
    DIVIDE([SGA], [Revenue]) * 100,
    BLANK()
)

-- البحث والتطوير / Research & Development
RD = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "6300",
    Financials[Statement Type], "IS"
)

-- الاستهلاك والإطفاء / Depreciation & Amortization
DA = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], {"6400", "6500"},
    Financials[Statement Type], "IS"
)
```

### 1.5 الدخل التشغيلي | Operating Income

```dax
-- الدخل التشغيلي / EBIT (Earnings Before Interest & Tax)
EBIT = [Gross_Profit] - [OpEx]

-- EBIT - السنة الماضية / Prior Year EBIT
PY_EBIT = 
CALCULATE(
    [EBIT],
    DATEADD(Dates[Date], -12, MONTH)
)

-- EBITDA (الدخل قبل الفوائد والضرائب والاستهلاك والإطفاء)
EBITDA = [EBIT] + [DA]

-- EBITDA - السنة الماضية / Prior Year EBITDA
PY_EBITDA = 
CALCULATE(
    [EBITDA],
    DATEADD(Dates[Date], -12, MONTH)
)

-- هامش EBITDA / EBITDA Margin %
EBITDA_Margin_Pct = 
IFERROR(
    DIVIDE([EBITDA], [Revenue]) * 100,
    BLANK()
)

-- هامش التشغيل / Operating Margin %
Op_Margin_Pct = 
IFERROR(
    DIVIDE([EBIT], [Revenue]) * 100,
    BLANK()
)
```

### 1.6 البنود غير التشغيلية | Non-Operating Items

```dax
-- مصاريف الفوائد / Interest Expense
Interest_Expense = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "7100",
    Financials[Statement Type], "IS"
)

-- إيرادات الفوائد / Interest Income
Interest_Income = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "7000",
    Financials[Statement Type], "IS"
)

-- صافي الفوائد / Net Interest
Net_Interest = [Interest_Income] - [Interest_Expense]

-- أرباح/(خسائر) العملات الأجنبية / Foreign Exchange Gain / (Loss)
FX_Gain_Loss = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "7200",
    Financials[Statement Type], "IS"
)

-- إجمالي البنود غير التشغيلية / Total Non-Operating Items
Non_Op_Items = [Net_Interest] + [FX_Gain_Loss]
```

### 1.7 الضرائب والدخل النهائي | Taxes & Final Income

```dax
-- الدخل قبل الضرائب / Income Before Tax (EBT)
Income_Before_Tax = [EBIT] + [Non_Op_Items]

-- مصروف ضريبة الدخل / Income Tax Expense
Tax_Expense = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "7200",
    Financials[Statement Type], "IS"
)

-- صافي الدخل / Net Income
Net_Income = [Income_Before_Tax] - [Tax_Expense]

-- صافي الدخل - السنة الماضية / Prior Year Net Income
PY_Net_Income = 
CALCULATE(
    [Net_Income],
    DATEADD(Dates[Date], -12, MONTH)
)

-- هامش صافي الربح / Net Profit Margin %
Net_Margin_Pct = 
IFERROR(
    DIVIDE([Net_Income], [Revenue]) * 100,
    BLANK()
)

-- معدل الضريبة الفعلي / Effective Tax Rate
Eff_Tax_Rate = 
IFERROR(
    DIVIDE([Tax_Expense], [Income_Before_Tax]) * 100,
    BLANK()
)
```

---

## 💰 الفئة 2: مقاييس الميزانية العمومية | SECTION 2: Balance Sheet Measures

### 2.1 الأصول | Assets

```dax
-- إجمالي الأصول المتداولة / Total Current Assets
Current_Assets = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Current Assets",
    Financials[Statement Type], "BS"
)

-- النقد والنقدية / Cash & Cash Equivalents
Cash = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "1000",
    Financials[Statement Type], "BS"
)

-- الأوراق المالية قصيرة الأجل / Short-Term Marketable Securities
ST_Securities = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "1100",
    Financials[Statement Type], "BS"
)

-- الذمم المدينة / Accounts Receivable
Receivables = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "1200",
    Financials[Statement Type], "BS"
)

-- المخزون / Inventory
Inventory = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "1300",
    Financials[Statement Type], "BS"
)

-- المصاريف المدفوعة مقدماً / Prepaid Expenses
Prepaid_Expenses = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "1400",
    Financials[Statement Type], "BS"
)

-- إجمالي الأصول غير المتداولة / Total Non-Current Assets
NonCurrent_Assets = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Non-Current Assets",
    Financials[Statement Type], "BS"
)

-- الأصول الثابتة الإجمالية / Gross Fixed Assets
Gross_Fixed_Assets = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "2000",
    Financials[Statement Type], "BS"
)

-- مجمع الاستهلاك / Accumulated Depreciation
Accumulated_Depr = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "2100",
    Financials[Statement Type], "BS"
)

-- صافي الأصول الثابتة / Net Fixed Assets
Net_Fixed_Assets = [Gross_Fixed_Assets] + [Accumulated_Depr]

-- الاستثمارات طويلة الأجل / Long-Term Investments
LT_Investments = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "2200",
    Financials[Statement Type], "BS"
)

-- إجمالي الأصول / Total Assets
Total_Assets = [Current_Assets] + [NonCurrent_Assets]
```

### 2.2 المطلوبات | Liabilities

```dax
-- إجمالي المطلوبات المتداولة / Total Current Liabilities
Current_Liabilities = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Current Liabilities",
    Financials[Statement Type], "BS"
)

-- الذمم الدائنة / Accounts Payable
Payables = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "3000",
    Financials[Statement Type], "BS"
)

-- القروض قصيرة الأجل / Short-Term Debt
ST_Debt = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "3100",
    Financials[Statement Type], "BS"
)

-- الجزء الجاري من الديون طويلة الأجل / Current Portion of LT Debt
Current_LT_Debt = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "3200",
    Financials[Statement Type], "BS"
)

-- المستحقات / Accrued Liabilities
Accrued_Liabilities = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "3300",
    Financials[Statement Type], "BS"
)

-- إجمالي المطلوبات طويلة الأجل / Total Long-Term Liabilities
LT_Liabilities = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Long-Term Liabilities",
    Financials[Statement Type], "BS"
)

-- الديون طويلة الأجل / Long-Term Debt
LT_Debt = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "3400",
    Financials[Statement Type], "BS"
)

-- إجمالي الديون / Total Debt
Total_Debt = [ST_Debt] + [Current_LT_Debt] + [LT_Debt]

-- صافي الديون / Net Debt
Net_Debt = [Total_Debt] - [Cash]

-- إجمالي المطلوبات / Total Liabilities
Total_Liabilities = [Current_Liabilities] + [LT_Liabilities]
```

### 2.3 حقوق الملكية | Equity

```dax
-- إجمالي حقوق الملكية / Total Equity
Total_Equity = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Equity",
    Financials[Statement Type], "BS"
)

-- رأس المال المدفوع / Share Capital
Share_Capital = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "4000",
    Financials[Statement Type], "BS"
)

-- علاوة إصدار الأسهم / Additional Paid-in Capital
Additional_Capital = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "4100",
    Financials[Statement Type], "BS"
)

-- الأرباح المحتجزة / Retained Earnings
Retained_Earnings = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "4200",
    Financials[Statement Type], "BS"
)

-- الدخل الشامل الآخر / Other Comprehensive Income
OCI = 
SUMIFS(
    Financials[Amount],
    Financials[Account Code], "4300",
    Financials[Statement Type], "BS"
)
```

---

## 💧 الفئة 3: مقاييس التدفق النقدي | SECTION 3: Cash Flow Measures

```dax
-- التدفق النقدي من الأنشطة التشغيلية / Operating Cash Flow (OCF)
OCF = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Operating Activities",
    Financials[Statement Type], "CF"
)

-- النفقات الرأسمالية / Capital Expenditures (CapEx)
CapEx = 
ABS(SUMIFS(
    Financials[Amount],
    Financials[Account Code], "5000",
    Financials[Statement Type], "CF"
))

-- التدفق النقدي الحر / Free Cash Flow (FCF)
FCF = [OCF] - [CapEx]

-- عائد التدفق النقدي الحر / FCF Yield
FCF_Yield = 
IFERROR(
    DIVIDE([FCF], [Revenue]) * 100,
    BLANK()
)

-- نسبة تحويل FCF / FCF Conversion Ratio
FCF_Conversion = 
IFERROR(
    DIVIDE([FCF], [Net_Income]) * 100,
    BLANK()
)

-- التدفق النقدي من الأنشطة الاستثمارية / Investing Cash Flow (ICF)
ICF = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Investing Activities",
    Financials[Statement Type], "CF"
)

-- التدفق النقدي من الأنشطة التمويلية / Financing Cash Flow (FCF)
Financing_CF = 
SUMIFS(
    Financials[Amount],
    Financials[Category], "Financing Activities",
    Financials[Statement Type], "CF"
)

-- صافي التغير في النقد / Net Change in Cash
Net_Cash_Change = [OCF] + [ICF] + [Financing_CF]
```

---

## 📈 الفئة 4: مقاييس النسب المالية - السيولة | SECTION 4: Liquidity Ratios

```dax
-- نسبة التداول / Current Ratio
Current_Ratio = 
IFERROR(
    DIVIDE([Current_Assets], [Current_Liabilities]),
    BLANK()
)

-- نسبة السيولة السريعة / Quick Ratio
Quick_Ratio = 
IFERROR(
    DIVIDE(
        [Current_Assets] - [Inventory],
        [Current_Liabilities]
    ),
    BLANK()
)

-- نسبة النقد / Cash Ratio
Cash_Ratio = 
IFERROR(
    DIVIDE([Cash], [Current_Liabilities]),
    BLANK()
)

-- نسبة التدفق التشغيلي / Operating Cash Flow Ratio
OCF_Ratio = 
IFERROR(
    DIVIDE([OCF], [Current_Liabilities]),
    BLANK()
)

-- رأس المال العامل / Working Capital (in Millions)
Working_Capital = [Current_Assets] - [Current_Liabilities]

-- نسبة رأس المال العامل / Working Capital Ratio
WC_Ratio = 
IFERROR(
    DIVIDE([Working_Capital], [Revenue]) * 100,
    BLANK()
)

-- حالة السيولة / Liquidity Status
Liquidity_Status = 
IF(
    [Current_Ratio] > 1.5,
    "🟢 HEALTHY",
    IF(
        [Current_Ratio] > 1.0,
        "🟡 CAUTION",
        "🔴 RISK"
    )
)
```

---

## 💹 الفئة 5: مقاييس النسب المالية - الملاءة | SECTION 5: Solvency Ratios

```dax
-- نسبة الدين للملكية / Debt-to-Equity Ratio
Debt_to_Equity = 
IFERROR(
    DIVIDE([Total_Debt], [Total_Equity]),
    BLANK()
)

-- نسبة الديون / Debt Ratio
Debt_Ratio = 
IFERROR(
    DIVIDE([Total_Liabilities], [Total_Assets]),
    BLANK()
)

-- صافي الدين إلى EBITDA / Net Debt to EBITDA
Net_Debt_to_EBITDA = 
IFERROR(
    DIVIDE([Net_Debt], [EBITDA]),
    BLANK()
)

-- نسبة تغطية الفوائد / Interest Coverage Ratio
Interest_Coverage = 
IFERROR(
    DIVIDE([EBIT], [Interest_Expense]),
    BLANK()
)

-- نسبة تغطية خدمة الدين / Debt Service Coverage Ratio
DSCR = 
IFERROR(
    DIVIDE(
        [OCF],
        [Interest_Expense] + [ST_Debt]
    ),
    BLANK()
)

-- حالة الملاءة / Solvency Status
Solvency_Status = 
IF(
    [Net_Debt_to_EBITDA] < 3,
    "🟢 HEALTHY",
    IF(
        [Net_Debt_to_EBITDA] < 4,
        "🟡 CAUTION",
        "🔴 RISK"
    )
)
```

---

## 🎯 الفئة 6: مقاييس النسب المالية - الربحية | SECTION 6: Profitability Ratios

```dax
-- العائد على الأصول / Return on Assets (ROA)
ROA = 
IFERROR(
    DIVIDE([Net_Income], [Total_Assets]) * 100,
    BLANK()
)

-- العائد على حقوق الملكية / Return on Equity (ROE)
ROE = 
IFERROR(
    DIVIDE([Net_Income], [Total_Equity]) * 100,
    BLANK()
)

-- العائد على رأس المال المستثمر / Return on Invested Capital (ROIC)
ROIC = 
IFERROR(
    DIVIDE(
        [EBIT] * (1 - 0.25),
        [Total_Equity] + [Total_Debt]
    ) * 100,
    BLANK()
)

-- حالة الربحية / Profitability Status
Profitability_Status = 
IF(
    [Net_Margin_Pct] > 10,
    "🟢 STRONG",
    IF(
        [Net_Margin_Pct] > 5,
        "🟡 MODERATE",
        "🔴 WEAK"
    )
)
```

---

## ⚙️ الفئة 7: مقاييس النسب المالية - الكفاءة | SECTION 7: Efficiency Ratios

```dax
-- معدل دوران الأصول / Asset Turnover
Asset_Turnover = 
IFERROR(
    DIVIDE([Revenue], [Total_Assets]),
    BLANK()
)

-- معدل دوران المخزون / Inventory Turnover
Inventory_Turnover = 
IFERROR(
    DIVIDE([COGS], [Inventory]),
    BLANK()
)

-- أيام المخزون المعلقة / Days Inventory Outstanding (DIO)
DIO = 
IFERROR(
    DIVIDE(365, [Inventory_Turnover]),
    BLANK()
)

-- معدل دوران الذمم المدينة / Receivables Turnover
Receivables_Turnover = 
IFERROR(
    DIVIDE([Revenue], [Receivables]),
    BLANK()
)

-- أيام المبيعات المعلقة / Days Sales Outstanding (DSO)
DSO = 
IFERROR(
    DIVIDE(365, [Receivables_Turnover]),
    BLANK()
)

-- معدل دوران الذمم الدائنة / Payables Turnover
Payables_Turnover = 
IFERROR(
    DIVIDE([COGS], [Payables]),
    BLANK()
)

-- أيام الذمم الدائنة / Days Payable Outstanding (DPO)
DPO = 
IFERROR(
    DIVIDE(365, [Payables_Turnover]),
    BLANK()
)

-- دورة تحويل النقد / Cash Conversion Cycle (CCC)
CCC = [DIO] + [DSO] - [DPO]
```

---

## 📊 الفئة 8: مقاييس المقارنة والنمو | SECTION 8: Growth & Comparison Measures

```dax
-- النمو شهري (MoM) / Month-over-Month Growth %
MoM_Growth_Pct = 
IFERROR(
    DIVIDE(
        [Revenue] - [Prior_Revenue],
        [Prior_Revenue]
    ) * 100,
    BLANK()
)

-- النمو ربعي (QoQ) / Quarter-over-Quarter Growth %
QoQ_Growth_Pct = 
IFERROR(
    DIVIDE(
        [Revenue] - CALCULATE([Revenue], DATEADD(Dates[Date], -3, MONTH)),
        CALCULATE([Revenue], DATEADD(Dates[Date], -3, MONTH))
    ) * 100,
    BLANK()
)

-- النمو السنوي (YoY) / Year-over-Year Growth %
YoY_Growth_Pct = 
IFERROR(
    DIVIDE(
        [Revenue] - [PY_Revenue],
        [PY_Revenue]
    ) * 100,
    BLANK()
)

-- نسبة التغير في EBITDA / EBITDA Change %
EBITDA_Change_Pct = 
IFERROR(
    DIVIDE(
        [EBITDA] - [PY_EBITDA],
        [PY_EBITDA]
    ) * 100,
    BLANK()
)

-- نسبة التغير في صافي الدخل / Net Income Change %
NI_Change_Pct = 
IFERROR(
    DIVIDE(
        [Net_Income] - [PY_Net_Income],
        [PY_Net_Income]
    ) * 100,
    BLANK()
)
```

---

## 📋 الفئة 9: مقاييس الانحرافات والميزانية | SECTION 9: Variance & Budget Measures

```dax
-- الانحراف المطلق / Variance Amount
Variance_Amount = [Actual_Value] - [Budget_Value]

-- الانحراف النسبي / Variance %
Variance_Pct = 
IFERROR(
    DIVIDE(
        [Variance_Amount],
        [Budget_Value]
    ) * 100,
    BLANK()
)

-- حالة الانحراف / Variance Status
Variance_Status = 
IF(
    [Variance_Pct] > 5,
    "🔴 UNFAVORABLE",
    IF(
        [Variance_Pct] < -5,
        "🟢 FAVORABLE",
        "🟡 ON TRACK"
    )
)

-- استخدام الميزانية / Budget Utilization %
Budget_Utilization_Pct = 
IFERROR(
    DIVIDE(
        [Actual_Value],
        [Budget_Value]
    ) * 100,
    BLANK()
)

-- حالة استخدام الميزانية / Budget Utilization Status
Budget_Status = 
IF(
    [Budget_Utilization_Pct] > 105,
    "🔴 OVER BUDGET",
    IF(
        [Budget_Utilization_Pct] > 90,
        "🟡 APPROACHING LIMIT",
        "🟢 WITHIN BUDGET"
    )
)
```

---

## 🔍 الفئة 10: مقاييس متقدمة | SECTION 10: Advanced Measures

### 10.1 المتوسطات المتحركة | Moving Averages

```dax
-- متوسط متحرك 3 أشهر / 3-Month Moving Average
MA_3M_Revenue = 
AVERAGEX(
    DATESBETWEEN(
        Dates[Date],
        DATEADD(MAX(Dates[Date]), -2, MONTH),
        MAX(Dates[Date])
    ),
    [Revenue]
)

-- متوسط متحرك 12 شهر / 12-Month Moving Average
MA_12M_Revenue = 
AVERAGEX(
    DATESBETWEEN(
        Dates[Date],
        DATEADD(MAX(Dates[Date]), -11, MONTH),
        MAX(Dates[Date])
    ),
    [Revenue]
)
```

### 10.2 معدل النمو السنوي المركب | CAGR

```dax
-- معدل النمو السنوي المركب 5 سنوات / 5-Year CAGR
CAGR_5Y = 
IFERROR(
    POWER(
        DIVIDE(
            CALCULATE([Revenue], DATEADD(Dates[Date], 0, YEAR)),
            CALCULATE([Revenue], DATEADD(Dates[Date], -60, MONTH))
        ),
        1/5
    ) - 1 * 100,
    BLANK()
)
```

### 10.3 مقاييس الترتيب | Ranking Measures

```dax
-- رتبة الإيرادات حسب الفئة / Revenue Rank by Category
Revenue_Rank = 
RANKX(
    ALL(Accounts[Category]),
    [Revenue],
    ,
    DESC
)

-- النسبة المئوية لإجمالي الإيرادات / % of Total Revenue
Pct_of_Total_Revenue = 
DIVIDE(
    [Revenue],
    CALCULATE(
        [Revenue],
        ALL(Accounts[Category])
    )
) * 100
```

---

## 🏥 الفئة 11: مؤشر ألتمان Z-Score | SECTION 11: Altman Z-Score

```dax
-- Altman Z-Score (للشركات الصناعية / For Manufacturing Companies)
Altman_Z_Score = 
1.2 * DIVIDE([Working_Capital], [Total_Assets]) +
1.4 * DIVIDE([Retained_Earnings], [Total_Assets]) +
3.3 * DIVIDE([EBIT], [Total_Assets]) +
0.6 * 1.0 +  -- Market Cap Ratio (تقدير / Estimate = 1.0)
0.999 * DIVIDE([Revenue], [Total_Assets])

-- حالة Z-Score / Z-Score Status
Z_Score_Status = 
IF(
    [Altman_Z_Score] > 2.99,
    "🟢 SAFE ZONE",
    IF(
        [Altman_Z_Score] > 1.81,
        "🟡 GREY ZONE",
        "🔴 DISTRESS ZONE"
    )
)
```

---

## 🎛️ الفئة 12: مقاييس الحالة والتنبيهات | SECTION 12: Status & Alert Measures

```dax
-- مؤشر الصحة المالية العامة / Overall Health Score
Health_Score = 
(
    IF([Liquidity_Status] = "🟢 HEALTHY", 3, IF([Liquidity_Status] = "🟡 CAUTION", 2, 1)) +
    IF([Solvency_Status] = "🟢 HEALTHY", 3, IF([Solvency_Status] = "🟡 CAUTION", 2, 1)) +
    IF([Profitability_Status] = "🟢 STRONG", 3, IF([Profitability_Status] = "🟡 MODERATE", 2, 1))
) / 9 * 100

-- حالة الصحة العامة / Health Status
Health_Score_Status = 
IF(
    [Health_Score] >= 75,
    "🟢 EXCELLENT",
    IF(
        [Health_Score] >= 50,
        "🟡 GOOD",
        IF(
            [Health_Score] >= 25,
            "🟠 FAIR",
            "🔴 POOR"
        )
    )
)

-- تعليق تنبيه / Alert Comment
Alert_Comment = 
SWITCH(
    TRUE(),
    [Current_Ratio] < 1.0, "⚠️ Low liquidity - Current Ratio < 1.0",
    [Net_Debt_to_EBITDA] > 4, "⚠️ High leverage - Net Debt/EBITDA > 4x",
    [Interest_Coverage] < 2, "⚠️ Weak interest coverage < 2x",
    [Gross_Margin_Pct] < 20, "⚠️ Declining gross margin < 20%",
    "✅ All metrics within healthy range"
)
```

---

## 📚 جدول المرجع السريع | Quick Reference Table

| المقياس | الصيغة المختصرة | الاستخدام | المعيار الصحي |
|--------|-------------|---------|------------|
| **Revenue** | SUM | إجمالي الإيرادات | ↑ مرتفع |
| **Gross Profit** | Rev - COGS | الربح الأساسي | ↑ مرتفع |
| **EBITDA** | EBIT + D&A | قبل الفوائد والضرائب | ↑ مرتفع |
| **Net Income** | ... - Tax | الربح النهائي | ↑ مرتفع |
| **Current Ratio** | CA / CL | السيولة قصيرة الأجل | 1.5 - 3.0 |
| **Quick Ratio** | (CA-Inv) / CL | السيولة الفورية | 1.0 - 2.0 |
| **Debt-to-Equity** | Debt / Equity | مستوى الرفع | < 2.0 |
| **ROE** | NI / Equity | عائد الملكية | > 15% |
| **ROA** | NI / Assets | عائد الأصول | > 5% |
| **FCF** | OCF - CapEx | النقد الحر | ↑ مرتفع |
| **DSO** | 365 / Rev Turnover | أيام المبيعات | ↓ منخفض |
| **CCC** | DIO + DSO - DPO | دورة النقد | ↓ منخفض |

---

## 🔧 نصائح مهمة | Important Tips

```
✅ استخدم IFERROR أو DIVIDE مع معامل ثالث لتجنب الأخطاء
✅ اختبر جميع المقاييس مع بيانات فعلية
✅ أضف تعليقات واضحة للمقاييس المعقدة
✅ استخدم نفس الوحدات في جميع الحسابات (ملايين / آلاف)
✅ حدّث البيانات بانتظام واختبر المقاييس
```

---

**آخر تحديث:** 2025-06-08 | **الإصدار:** 1.0 | **المقاييس:** 80+ Measures

