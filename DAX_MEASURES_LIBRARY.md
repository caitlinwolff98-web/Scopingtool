# DAX Measures Library
## Complete Reference for Bidvest ISA 600 Scoping Tool

**Version:** 5.0 Production Ready
**Last Updated:** November 2025
**Power BI Compatibility:** Power BI Desktop (all versions)

---

## 📖 Table of Contents

1. [Quick Start](#quick-start)
2. [Basic Count Measures](#basic-count-measures)
3. [Scoping Status Measures](#scoping-status-measures)
4. [Coverage Percentage Measures](#coverage-percentage-measures)
5. [Amount-Based Measures](#amount-based-measures)
6. [FSLI-Specific Measures](#fsli-specific-measures)
7. [Division-Based Measures](#division-based-measures)
8. [Threshold Analysis Measures](#threshold-analysis-measures)
9. [Advanced Analytical Measures](#advanced-analytical-measures)
10. [Time Intelligence Measures](#time-intelligence-measures)
11. [Troubleshooting](#troubleshooting)

---

## 🚀 Quick Start

### How to Add DAX Measures in Power BI

1. Open your Power BI file
2. Click **Data view** (left sidebar - table icon)
3. Select the table where you want the measure (usually **Scoping_Control_Table**)
4. Click **New Measure** in the ribbon
5. Copy and paste the DAX code below
6. Press **Enter** to create the measure
7. Format the measure (percentage, currency, etc.) in the Modeling ribbon

### Recommended Measures for New Users

Start with these 5 essential measures:

1. **Total Packs** - Count of all packs
2. **Scoped In Packs** - Count of packs scoped in
3. **Coverage %** - Percentage of packs scoped in
4. **Total Amount Scoped In** - Sum of amounts scoped in
5. **Coverage % by Amount** - Percentage of total amount scoped in

Then add specialized measures as needed.

---

## 1. Basic Count Measures

### Total Packs

**Purpose:** Count all packs (excludes consolidated entity)

```DAX
Total Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Display in Card visual for dashboard summary
- Use as denominator in coverage calculations
- Filter by Division to get division-specific counts

**Expected Result:** Integer (e.g., 45 packs)

---

### Total FSLIs

**Purpose:** Count all Financial Statement Line Items

```DAX
Total FSLIs =
DISTINCTCOUNT(Scoping_Control_Table[FSLI])
```

**Usage:**
- Display total number of line items being analyzed
- Compare across divisions or periods
- Validate data completeness

**Expected Result:** Integer (e.g., 250 FSLIs)

---

### Total Divisions

**Purpose:** Count active divisions (excludes "Not Categorized")

```DAX
Total Divisions =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Division]),
    Scoping_Control_Table[Division] <> "Not Categorized"
)
```

**Usage:**
- Summary statistics
- Division-level analysis
- ISA 600 component reporting

**Expected Result:** Integer (e.g., 8 divisions)

---

### Total Records

**Purpose:** Count all scoping decision records

```DAX
Total Records =
COUNTROWS(Scoping_Control_Table)
```

**Usage:**
- Data validation
- Performance monitoring
- Audit trail verification

**Expected Result:** Large integer (e.g., 11,250 records = 45 packs × 250 FSLIs)

---

## 2. Scoping Status Measures

### Scoped In Packs (Auto)

**Purpose:** Count packs automatically scoped in by thresholds

```DAX
Scoped In Packs (Auto) =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Scoped In (Auto)",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Track effectiveness of threshold-based scoping
- Separate automatic vs. manual decisions
- Audit trail for scoping methodology

**Expected Result:** Integer (e.g., 12 packs)

---

### Scoped In Packs (Manual)

**Purpose:** Count packs manually scoped in by auditor

```DAX
Scoped In Packs (Manual) =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Scoped In (Manual)",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Track manual scoping decisions
- Professional judgment documentation
- Compare with automatic scoping

**Expected Result:** Integer (e.g., 8 packs)

---

### Scoped In Packs (Total)

**Purpose:** Count all packs scoped in (auto + manual)

```DAX
Scoped In Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Primary scoping coverage metric
- Dashboard summary cards
- ISA 600 reporting

**Expected Result:** Integer (e.g., 20 packs)

---

### Not Scoped Packs

**Purpose:** Count packs not yet scoped in

```DAX
Not Scoped Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Not Scoped",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Identify remaining scoping work
- Gap analysis
- Coverage planning

**Expected Result:** Integer (e.g., 25 packs)

---

### Scoped Out Packs

**Purpose:** Count packs explicitly excluded from scope

```DAX
Scoped Out Packs =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Scoping Status] = "Scoped Out",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Track exclusion decisions
- Risk assessment
- Scope reduction documentation

**Expected Result:** Integer (e.g., 0 packs)

---

## 3. Coverage Percentage Measures

### Coverage % (by Pack Count)

**Purpose:** Percentage of packs scoped in

```DAX
Coverage % =
DIVIDE(
    [Scoped In Packs],
    [Total Packs],
    0
)
```

**Format:** Percentage with 1 decimal place (e.g., 44.4%)

**Usage:**
- Primary KPI for scoping progress
- Dashboard headline metric
- ISA 600 documentation

**Interpretation:**
- < 50%: Low coverage, more scoping needed
- 50-80%: Moderate coverage, targeted scoping
- > 80%: High coverage, review for efficiency

---

### Coverage % (by Amount)

**Purpose:** Percentage of total amounts scoped in

```DAX
Coverage % by Amount =
DIVIDE(
    [Total Amount Scoped In],
    [Total Amount (All Packs)],
    0
)
```

**Format:** Percentage with 1 decimal place

**Usage:**
- Materiality-based coverage assessment
- More meaningful than pack count for ISA 600
- Risk-based scoping validation

**Expected Result:** Usually higher than pack count % (e.g., 85% vs. 44%)

---

### Untested % (by Pack Count)

**Purpose:** Percentage of packs not scoped in

```DAX
Untested % =
DIVIDE(
    [Not Scoped Packs],
    [Total Packs],
    0
)
```

**Format:** Percentage with 1 decimal place

**Usage:**
- Gap analysis
- Risk assessment
- Planning remaining audit work

**Formula Check:** Should equal `1 - Coverage %`

---

### Untested % (by Amount)

**Purpose:** Percentage of amounts not scoped in

```DAX
Untested % by Amount =
DIVIDE(
    [Total Amount Not Scoped],
    [Total Amount (All Packs)],
    0
)
```

**Format:** Percentage with 1 decimal place

**Usage:**
- Materiality gap analysis
- Risk exposure assessment
- ISA 600 component coverage

---

### Coverage % per FSLI

**Purpose:** Coverage percentage for each Financial Statement Line Item

```DAX
Coverage % per FSLI =
VAR ScopedAmount =
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"}
    )
VAR TotalAmount =
    SUM(Scoping_Control_Table[Amount])
RETURN
    DIVIDE(ScopedAmount, TotalAmount, 0)
```

**Format:** Percentage with 1 decimal place

**Usage:**
- FSLI-level coverage analysis
- Identify under-scoped line items
- Targeted scoping decisions

**Context:** Automatically filters to selected FSLI when used in visual

---

### Coverage % per Division

**Purpose:** Coverage percentage for each division

```DAX
Coverage % per Division =
VAR ScopedPacks =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalPacks =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Is Consolidated] = "No"
    )
RETURN
    DIVIDE(ScopedPacks, TotalPacks, 0)
```

**Format:** Percentage with 1 decimal place

**Usage:**
- Division-level scoping analysis
- ISA 600 component audit
- Resource allocation

---

## 4. Amount-Based Measures

### Total Amount (All Packs)

**Purpose:** Sum of all amounts in scoping analysis

```DAX
Total Amount (All Packs) =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency with thousands separator (e.g., R 1,250,000,000)

**Usage:**
- Baseline for coverage calculations
- Materiality assessments
- Threshold setting

---

### Total Amount Scoped In

**Purpose:** Sum of amounts for scoped-in packs

```DAX
Total Amount Scoped In =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency with thousands separator

**Usage:**
- Coverage assessment by value
- Materiality coverage tracking
- ISA 600 reporting

---

### Total Amount Scoped In (Auto)

**Purpose:** Sum of amounts automatically scoped in

```DAX
Total Amount Scoped In (Auto) =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] = "Scoped In (Auto)",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- Evaluate threshold effectiveness
- Automatic scoping impact
- Audit methodology documentation

---

### Total Amount Scoped In (Manual)

**Purpose:** Sum of amounts manually scoped in

```DAX
Total Amount Scoped In (Manual) =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] = "Scoped In (Manual)",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- Professional judgment impact
- Manual scoping effectiveness
- Incremental coverage from manual work

---

### Total Amount Not Scoped

**Purpose:** Sum of amounts not yet scoped in

```DAX
Total Amount Not Scoped =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] = "Not Scoped",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- Gap analysis
- Risk exposure
- Remaining scoping opportunity

---

### Average Amount per Pack

**Purpose:** Average amount per pack

```DAX
Average Amount per Pack =
DIVIDE(
    [Total Amount (All Packs)],
    [Total Packs],
    0
)
```

**Format:** Currency

**Usage:**
- Identify significant packs
- Threshold setting guidance
- Pack prioritization

---

## 5. FSLI-Specific Measures

### Packs with Selected FSLI

**Purpose:** Count packs that have amounts in the filtered FSLI

```DAX
Packs with Selected FSLI =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Amount] <> 0,
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- FSLI-specific analysis
- Identify which packs have activity in specific line items
- Targeted scoping

**Context Sensitive:** Filters to selected FSLI automatically

---

### Amount for Selected FSLI

**Purpose:** Total amount for filtered FSLI

```DAX
Amount for Selected FSLI =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- FSLI materiality assessment
- Threshold validation
- Line item significance

---

### Scoped In Amount for Selected FSLI

**Purpose:** Scoped-in amount for filtered FSLI

```DAX
Scoped In Amount for Selected FSLI =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- FSLI-specific coverage assessment
- Line item scoping effectiveness
- Targeted gap analysis

---

### Top FSLI by Amount

**Purpose:** Rank FSLIs by total amount

```DAX
FSLI Rank by Amount =
RANKX(
    ALL(Scoping_Control_Table[FSLI]),
    [Amount for Selected FSLI],
    ,
    DESC,
    DENSE
)
```

**Format:** Whole number (1, 2, 3...)

**Usage:**
- Identify material FSLIs
- Prioritize scoping effort
- Risk-based selection

**Tip:** Use in table visual with TOPN filter to show top 10 FSLIs

---

## 6. Division-Based Measures

### Packs in Selected Division

**Purpose:** Count packs in filtered division

```DAX
Packs in Selected Division =
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Division-level analysis
- Component sizing (ISA 600)
- Resource allocation

**Context Sensitive:** Filters to selected division automatically

---

### Amount in Selected Division

**Purpose:** Total amount in filtered division

```DAX
Amount in Selected Division =
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- Division materiality
- Component significance (ISA 600)
- Division-level scoping

---

### Division Coverage %

**Purpose:** Coverage for filtered division

```DAX
Division Coverage % =
VAR ScopedInDivision =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"},
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalInDivision =
    [Packs in Selected Division]
RETURN
    DIVIDE(ScopedInDivision, TotalInDivision, 0)
```

**Format:** Percentage

**Usage:**
- Division-specific coverage tracking
- Component audit progress
- ISA 600 reporting

---

### Division as % of Total

**Purpose:** Division's proportion of total amounts

```DAX
Division as % of Total =
DIVIDE(
    [Amount in Selected Division],
    CALCULATE(
        SUM(Scoping_Control_Table[Amount]),
        ALL(Scoping_Control_Table[Division]),
        Scoping_Control_Table[Is Consolidated] = "No"
    ),
    0
)
```

**Format:** Percentage

**Usage:**
- Identify significant components
- ISA 600 component classification
- Materiality allocation

---

## 7. Threshold Analysis Measures

### Packs Above Threshold (Dynamic)

**Purpose:** Count packs where selected FSLI exceeds threshold

**Note:** Requires threshold parameter (use What-If parameter or manual entry)

```DAX
Packs Above Threshold =
VAR ThresholdValue = 300000000 // Adjust this value
RETURN
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Amount] >= ThresholdValue,
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Threshold sensitivity analysis
- Optimize threshold settings
- Coverage vs. efficiency trade-off

**Customization:** Replace `300000000` with your threshold or create What-If parameter

---

### Coverage % at Threshold

**Purpose:** Coverage percentage if threshold is applied

```DAX
Coverage % at Threshold =
VAR ThresholdValue = 300000000 // Adjust this value
VAR PacksAbove =
    CALCULATE(
        DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
        Scoping_Control_Table[Amount] >= ThresholdValue,
        Scoping_Control_Table[Is Consolidated] = "No"
    )
VAR TotalPacks = [Total Packs]
RETURN
    DIVIDE(PacksAbove, TotalPacks, 0)
```

**Format:** Percentage

**Usage:**
- Test different threshold levels
- Optimize scoping strategy
- Balance coverage vs. effort

---

### Amount Above Threshold

**Purpose:** Total amount for packs exceeding threshold

```DAX
Amount Above Threshold =
VAR ThresholdValue = 300000000 // Adjust this value
RETURN
CALCULATE(
    SUM(Scoping_Control_Table[Amount]),
    Scoping_Control_Table[Amount] >= ThresholdValue,
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Format:** Currency

**Usage:**
- Materiality coverage at threshold
- Validate threshold effectiveness
- Risk assessment

---

## 8. Advanced Analytical Measures

### Scoping Efficiency Ratio

**Purpose:** Ratio of coverage % by amount to coverage % by pack count

```DAX
Scoping Efficiency Ratio =
DIVIDE(
    [Coverage % by Amount],
    [Coverage %],
    0
)
```

**Format:** Decimal (e.g., 1.91)

**Usage:**
- Assess scoping effectiveness
- High ratio (>1.5) = efficient scoping (covering large packs)
- Low ratio (<1.2) = less efficient (covering many small packs)

**Interpretation:**
- 1.0 = Proportional (pack sizes similar)
- > 1.5 = Efficient (focusing on large packs)
- < 1.0 = Error (should not occur)

---

### Incremental Coverage from Manual Scoping

**Purpose:** Additional coverage gained from manual scoping

```DAX
Incremental Coverage from Manual =
DIVIDE(
    [Scoped In Packs (Manual)],
    [Total Packs],
    0
)
```

**Format:** Percentage

**Usage:**
- Quantify value of manual professional judgment
- Evaluate auto vs. manual scoping split
- Resource allocation decisions

---

### Packs Needing Review

**Purpose:** Count packs close to threshold requiring judgment

**Requires:** Threshold parameter

```DAX
Packs Needing Review =
VAR ThresholdValue = 300000000
VAR BufferPercent = 0.10 // 10% buffer
VAR LowerBound = ThresholdValue * (1 - BufferPercent)
VAR UpperBound = ThresholdValue * (1 + BufferPercent)
RETURN
CALCULATE(
    DISTINCTCOUNT(Scoping_Control_Table[Pack Code]),
    Scoping_Control_Table[Amount] >= LowerBound,
    Scoping_Control_Table[Amount] <= UpperBound,
    Scoping_Control_Table[Scoping Status] = "Not Scoped",
    Scoping_Control_Table[Is Consolidated] = "No"
)
```

**Usage:**
- Identify borderline cases
- Focus manual review efforts
- Professional judgment documentation

---

### Coverage Gap

**Purpose:** Number of additional packs needed to reach target coverage

```DAX
Coverage Gap to 80% =
VAR TargetCoverage = 0.80
VAR TargetPacks = [Total Packs] * TargetCoverage
VAR CurrentScoped = [Scoped In Packs]
RETURN
    MAX(0, ROUND(TargetPacks - CurrentScoped, 0))
```

**Format:** Whole number

**Usage:**
- Planning remaining scoping work
- Resource allocation
- Progress tracking

**Customization:** Change `0.80` to your target coverage percentage

---

## 9. Time Intelligence Measures

**Note:** These measures require a Date table with proper relationships

### Coverage % Prior Period

**Purpose:** Compare current coverage to previous period

```DAX
Coverage % Prior Period =
CALCULATE(
    [Coverage %],
    DATEADD('Date'[Date], -1, MONTH)
)
```

**Format:** Percentage

**Usage:**
- Track scoping progress over time
- Month-over-month comparison
- Trend analysis

---

### Coverage % Change

**Purpose:** Change in coverage from prior period

```DAX
Coverage % Change =
[Coverage %] - [Coverage % Prior Period]
```

**Format:** Percentage (with +/- sign)

**Usage:**
- Progress monitoring
- Performance tracking
- Management reporting

---

## 10. Troubleshooting

### Common Issues and Solutions

#### Measure Returns BLANK

**Possible Causes:**
- Division by zero (use `DIVIDE` function with third parameter)
- Missing filter context
- No data in filtered context

**Solution:**
```DAX
// Instead of:
Coverage % = [Scoped In Packs] / [Total Packs]

// Use:
Coverage % = DIVIDE([Scoped In Packs], [Total Packs], 0)
```

#### Measure Counts Too High

**Possible Causes:**
- Not filtering out consolidated entity
- Not using DISTINCTCOUNT
- Incorrect filter context

**Solution:**
```DAX
// Always include:
Scoping_Control_Table[Is Consolidated] = "No"

// And use DISTINCTCOUNT for pack counts:
DISTINCTCOUNT(Scoping_Control_Table[Pack Code])
```

#### Coverage % Greater Than 100%

**Possible Causes:**
- Consolidated entity not excluded
- Duplicate records in data
- Incorrect numerator/denominator

**Solution:**
1. Verify consolidated entity is marked "Is Consolidated = Yes"
2. Check for duplicate Pack Code records
3. Verify both numerator and denominator use same filters

#### Measure Not Responding to Slicers

**Possible Causes:**
- Using ALL() incorrectly
- Filter context removed
- Relationship issues

**Solution:**
- Remove unnecessary ALL() functions
- Check table relationships in Model view
- Use ALLSELECTED() instead of ALL() to respect slicers

---

## 📚 Appendix: Measure Categories

### Essential Measures (Start Here)
1. Total Packs
2. Scoped In Packs
3. Coverage %
4. Total Amount Scoped In
5. Coverage % by Amount

### Core Scoping Measures
6. Scoped In Packs (Auto)
7. Scoped In Packs (Manual)
8. Not Scoped Packs
9. Untested %
10. Untested % by Amount

### Amount Analysis
11. Total Amount (All Packs)
12. Total Amount Not Scoped
13. Average Amount per Pack

### FSLI Analysis
14. Coverage % per FSLI
15. Packs with Selected FSLI
16. Amount for Selected FSLI

### Division Analysis
17. Coverage % per Division
18. Packs in Selected Division
19. Amount in Selected Division
20. Division as % of Total

### Advanced Analysis
21. Scoping Efficiency Ratio
22. Incremental Coverage from Manual
23. Packs Needing Review
24. Coverage Gap

---

## 🎯 Quick Reference Card

### Copy This to Create All Essential Measures

```DAX
// 1. Basic Counts
Total Packs = CALCULATE(DISTINCTCOUNT(Scoping_Control_Table[Pack Code]), Scoping_Control_Table[Is Consolidated] = "No")

Scoped In Packs = CALCULATE(DISTINCTCOUNT(Scoping_Control_Table[Pack Code]), Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"}, Scoping_Control_Table[Is Consolidated] = "No")

// 2. Coverage Percentages
Coverage % = DIVIDE([Scoped In Packs], [Total Packs], 0)

Coverage % by Amount = DIVIDE([Total Amount Scoped In], [Total Amount (All Packs)], 0)

// 3. Amounts
Total Amount (All Packs) = CALCULATE(SUM(Scoping_Control_Table[Amount]), Scoping_Control_Table[Is Consolidated] = "No")

Total Amount Scoped In = CALCULATE(SUM(Scoping_Control_Table[Amount]), Scoping_Control_Table[Scoping Status] IN {"Scoped In (Auto)", "Scoped In (Manual)"}, Scoping_Control_Table[Is Consolidated] = "No")
```

---

**Document Version:** 5.0
**Last Updated:** November 2025
**Total Measures:** 40+
**Difficulty Levels:** Beginner to Advanced

**Next Steps:**
- See [IMPLEMENTATION_GUIDE.md](IMPLEMENTATION_GUIDE.md) for setup instructions
- See [POWER_BI_EDIT_MODE_GUIDE.md](POWER_BI_EDIT_MODE_GUIDE.md) for manual scoping
- See [COMPREHENSIVE_GUIDE.md](COMPREHENSIVE_GUIDE.md) for complete technical reference

---

**Questions?** Check the Troubleshooting section or refer to Power BI DAX documentation at docs.microsoft.com/dax
