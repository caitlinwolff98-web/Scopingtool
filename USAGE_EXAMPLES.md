# Usage Examples - TGK Consolidation Scoping Tool

## Table of Contents
1. [Example 1: Basic Single-Segment Consolidation](#example-1-basic-single-segment-consolidation)
2. [Example 2: Multi-Segment with Discontinued Operations](#example-2-multi-segment-with-discontinued-operations)
3. [Example 3: Threshold-Based Scoping](#example-3-threshold-based-scoping)
4. [Example 4: Manual Scoping Workflow](#example-4-manual-scoping-workflow)
5. [Example 5: Hybrid Scoping Approach](#example-5-hybrid-scoping-approach)
6. [Example 6: Complex Multi-Division Analysis](#example-6-complex-multi-division-analysis)
7. [Common Scenarios](#common-scenarios)

---

## Example 1: Basic Single-Segment Consolidation

### Scenario
You have a simple consolidation workbook with UK entities only, no discontinued operations.

### Consolidation Workbook Structure
```
Workbook Name: UK_Consolidation_2024.xlsx

Tabs:
- TGK_UK
- TGK_Input_Continuing
- TGK_Journals
- TGK_Console
- Balance_Sheet
- Income_Statement
- Summary (not needed)
```

### Step-by-Step Process

#### 1. Open Files
```
✓ Open: UK_Consolidation_2024.xlsx
✓ Open: TGK_Scoping_Tool.xlsm
```

#### 2. Start Tool
- Click "Start TGK Scoping Tool" button
- Welcome dialog appears → Click **OK**

#### 3. Enter Workbook Name
```
Prompt: "Please enter the exact name of the TGK consolidation workbook"
Enter: UK_Consolidation_2024.xlsx
Click: OK
```

#### 4. Categorize Tabs
Tool displays pop-up dialogs to categorize each tab.

**Categorization Process:**

For each of the 7 tabs, a dialog will appear:

1. **Tab 1: TGK_UK**
   - Pop-up shows: "Tab 1 of 7 - Tab Name: TGK_UK"
   - Enter: `1` (for TGK Segment Tabs)
   - Next pop-up: "Enter the division name for this segment tab"
   - Enter: `UK`

2. **Tab 2: TGK_Input_Continuing**
   - Pop-up shows: "Tab 2 of 7 - Tab Name: TGK_Input_Continuing"
   - Enter: `3` (for TGK Input Continuing Operations Tab)

3. **Tab 3: TGK_Journals**
   - Pop-up shows: "Tab 3 of 7 - Tab Name: TGK_Journals"
   - Enter: `4` (for TGK Journals Continuing Tab)

4. **Tab 4: TGK_Console**
   - Pop-up shows: "Tab 4 of 7 - Tab Name: TGK_Console"
   - Enter: `5` (for TGK Console Continuing Tab)

5. **Tab 5: Balance_Sheet**
   - Pop-up shows: "Tab 5 of 7 - Tab Name: Balance_Sheet"
   - Enter: `6` (for TGK BS Tab)

6. **Tab 6: Income_Statement**
   - Pop-up shows: "Tab 6 of 7 - Tab Name: Income_Statement"
   - Enter: `7` (for TGK IS Tab)

7. **Tab 7: Summary**
   - Pop-up shows: "Tab 7 of 7 - Tab Name: Summary"
   - Enter: `9` (for Uncategorized)

**Result:** All tabs categorized via pop-up dialogs, validation passes

#### 5. Select Column Type
```
Prompt: "Which columns do you want to use?"
Options:
  - Consolidation/Consolidation Currency (25 columns)
  - Original/Entity Currency (25 columns)
  
Select: YES (for Consolidation/Consolidation Currency)
```

#### 6. Wait for Processing
```
Status bar shows:
"Processing Input Continuing tab..."
"Processing Journals tab..."
"Processing Console tab..."
"Creating FSLi Key Table..."
"Creating Pack Number Company Table..."
"Creating Percentage Tables..."
```

**Expected Time:** 1-2 minutes for small workbook

#### 7. Review Output
New workbook created: `Book1` (or similar)

**Tables Generated:**
- Control Panel (info sheet)
- Full Input Table (5 packs × 150 FSLis)
- Full Input Percentage
- Journals Table
- Journals Percentage
- Full Console Table
- Full Console Percentage
- FSLi Key Table
- Pack Number Company Table

#### 8. Save Output
```
File → Save As
Name: UK_Consolidation_2024_Output.xlsx
Type: Excel Workbook (*.xlsx)
```

### Expected Results

**Pack Number Company Table:**
| Pack Name | Pack Code | Division |
|-----------|-----------|----------|
| UK Holdco Ltd | BVT-001 | UK |
| UK Trading Ltd | BVT-002 | UK |
| UK Services Ltd | BVT-003 | UK |
| UK Property Ltd | BVT-004 | UK |
| UK Finance Ltd | BVT-005 | UK |

**Full Input Table Sample:**
| Pack | Revenue | Cost of Sales | Gross Profit | ... |
|------|---------|---------------|--------------|-----|
| UK Holdco Ltd | 50,000,000 | 0 | 50,000,000 | ... |
| UK Trading Ltd | 120,000,000 | 72,000,000 | 48,000,000 | ... |
| UK Services Ltd | 30,000,000 | 18,000,000 | 12,000,000 | ... |
| UK Property Ltd | 15,000,000 | 5,000,000 | 10,000,000 | ... |
| UK Finance Ltd | 5,000,000 | 1,000,000 | 4,000,000 | ... |

---

## Example 2: Multi-Segment with Discontinued Operations

### Scenario
Consolidation workbook with three geographic segments plus discontinued operations.

### Workbook Structure
```
Workbook Name: Global_Consolidation_Q4_2024.xlsx

Tabs:
- TGK_UK (15 entities)
- TGK_US (20 entities)
- TGK_Europe (10 entities)
- TGK_Discontinued (3 entities)
- TGK_Input_Continuing
- TGK_Journals_Continuing
- TGK_Console_Continuing
- Balance_Sheet
- Income_Statement
- Notes
- Cash_Flow
- Working_Paper_1
- Working_Paper_2
```

### Categorization Strategy

| Tab Name | Category | Division Name | Notes |
|----------|----------|---------------|-------|
| TGK_UK | TGK Segment Tabs | UK | 15 entities |
| TGK_US | TGK Segment Tabs | United States | 20 entities |
| TGK_Europe | TGK Segment Tabs | Europe | 10 entities |
| TGK_Discontinued | TGK Discontinued Opt Tab | | 3 entities |
| TGK_Input_Continuing | TGK Input Continuing Operations Tab | | Required |
| TGK_Journals_Continuing | TGK Journals Continuing Tab | | |
| TGK_Console_Continuing | TGK Console Continuing Tab | | |
| Balance_Sheet | TGK BS Tab | | |
| Income_Statement | TGK IS Tab | | |
| Notes | Uncategorized | | Not processed |
| Cash_Flow | Uncategorized | | Not processed |
| Working_Paper_1 | Pull Workings | | Optional |
| Working_Paper_2 | Pull Workings | | Optional |

### Special Considerations

**1. Division Name Prompts**
Tool will prompt for division names:
```
Prompt: "Please enter the division name for: TGK_UK"
Enter: UK
```
```
Prompt: "Please enter the division name for: TGK_US"
Enter: United States
```
```
Prompt: "Please enter the division name for: TGK_Europe"
Enter: Europe
```

**2. Uncategorized Tabs Warning**
```
Dialog: "The following tabs were not categorized:
- Notes
- Cash_Flow

These tabs will be ignored during processing.
Do you want to proceed?"

Select: YES
```

### Output Tables

**Pack Number Company Table (48 entities total):**
| Pack Name | Pack Code | Division |
|-----------|-----------|----------|
| UK Holdco Ltd | BVT-001 | UK |
| ... (14 more UK) | ... | UK |
| US Holdco Inc | BVT-016 | United States |
| ... (19 more US) | ... | United States |
| Europe Holdco SA | BVT-036 | Europe |
| ... (9 more Europe) | ... | Europe |
| Discontinued Ltd | BVT-046 | Discontinued |
| Discontinued Inc | BVT-047 | Discontinued |
| Discontinued SA | BVT-048 | Discontinued |

**Additional Tables:**
- Full Input Table (45 packs × 200 FSLis)
- Full Input Percentage
- Journals Table (45 packs × 200 FSLis)
- Journals Percentage
- Full Console Table (45 packs × 200 FSLis)
- Full Console Percentage
- Discontinued Table (3 packs × 200 FSLis)
- Discontinued Percentage
- FSLi Key Table (200 FSLis)
- Pack Number Company Table (48 packs)

**Processing Time:** 4-6 minutes

---

## Example 3: Threshold-Based Scoping

### Scenario
Using Power BI to automatically scope in entities based on Net Revenue threshold of $300M.

### Prerequisites
- Output workbook from Example 2 imported into Power BI
- Data transformed (unpivoted)
- Relationships created
- DAX measures configured

### Power BI Workflow

#### Step 1: Open Threshold Scoping Page
Navigate to "Threshold Scoping" report page

#### Step 2: Select FSLi
```
FSLi Selector: Net Revenue
```

#### Step 3: Set Threshold
```
Threshold Selector: $300,000,000
```

#### Step 4: Review Packs Meeting Threshold

**Packs Meeting Threshold Table:**
| Pack Name | FSLi | Amount | Meets Threshold |
|-----------|------|--------|-----------------|
| UK Trading Ltd | Net Revenue | $450,000,000 | Yes |
| US Operations Inc | Net Revenue | $520,000,000 | Yes |
| US Manufacturing Inc | Net Revenue | $380,000,000 | Yes |
| Europe Services SA | Net Revenue | $310,000,000 | Yes |

**Result:** 4 packs scoped in automatically

#### Step 5: Check Coverage

**Coverage Cards:**
```
Packs Scoped In: 4
Total Coverage %: 68.5%
Untested %: 31.5%
```

**Interpretation:**
- 4 out of 48 entities (8.3%)
- Represent 68.5% of total net revenue
- Efficient scope focused on largest entities

#### Step 6: Document

Export the "Packs Meeting Threshold" table:
```
Right-click table → Export Data → Excel
Save as: Threshold_Scoping_NetRevenue_300M.xlsx
```

### Adding Additional Threshold

#### Scenario: Also scope in based on Total Assets > $500M

**Steps:**
1. Change FSLi Selector to "Total Assets"
2. Set Threshold to $500,000,000
3. Review new packs meeting this threshold
4. Combined coverage now includes both revenue and asset thresholds

**Additional Packs:**
| Pack Name | FSLi | Amount | Meets Threshold |
|-----------|------|--------|-----------------|
| UK Property Ltd | Total Assets | $650,000,000 | Yes |
| US Real Estate Inc | Total Assets | $720,000,000 | Yes |

**Updated Coverage:**
```
Packs Scoped In: 6 (4 from revenue + 2 from assets)
Total Coverage %: 75.2%
Untested %: 24.8%
```

---

## Example 4: Manual Scoping Workflow

### Scenario
Manually select specific packs and FSLis that don't meet automatic thresholds but require testing due to audit risk.

### Power BI Workflow

#### Step 1: Navigate to Manual Selection Page

#### Step 2: Select Pack
```
Pack Picker: 
☑ Europe Manufacturing SA (below thresholds but high risk)
```

#### Step 3: Select Specific FSLis
```
FSLi Picker:
☑ Inventory (known reconciliation issues)
☑ Accounts Payable (new vendor system)
☑ Intercompany Receivables (complex transactions)
```

**Note:** Only these 3 FSLis for this pack, not all FSLis

#### Step 4: Review Selection

**Current Selection Table:**
| Pack Name | FSLi | Amount | Percentage |
|-----------|------|--------|------------|
| Europe Manufacturing SA | Inventory | $45,000,000 | 3.2% |
| Europe Manufacturing SA | Accounts Payable | $28,000,000 | 2.1% |
| Europe Manufacturing SA | Intercompany Receivables | $15,000,000 | 1.1% |

#### Step 5: Add More Packs

Repeat for other high-risk packs:
```
☑ UK Services Ltd → Trade Receivables, Goodwill
☑ US Tech Inc → Intangible Assets, Deferred Revenue
```

#### Step 6: Review Total Coverage

**Selection Summary:**
```
Packs Scoped In: 3
FSLis Scoped In: 8
Coverage Amount: $156,000,000
Coverage %: 6.5%
```

**Interpretation:**
- Manual selection adds 6.5% coverage
- Focused on specific risk areas
- Supplements threshold scoping

#### Step 7: Combine with Threshold

**Total Scope (Threshold + Manual):**
```
Threshold Scoping: 6 packs, 75.2% coverage
Manual Selection: 3 packs (8 FSLis), 6.5% coverage
Combined Total: 81.7% coverage
```

**Meets audit requirement:** Yes (target was 80%)

---

## Example 5: Hybrid Scoping Approach

### Scenario
Comprehensive scoping combining threshold, manual selection, and complete pack inclusion.

### Phase 1: Threshold Scoping

**Threshold 1 - Revenue Based:**
```
FSLi: Net Revenue
Threshold: $300,000,000
Result: 4 packs scoped in
Coverage: 68.5%
```

**Threshold 2 - Assets Based:**
```
FSLi: Total Assets
Threshold: $500,000,000
Result: 2 additional packs
Cumulative Coverage: 75.2%
```

### Phase 2: Manual FSLi Selection

**High-Risk Items:**
```
For packs NOT yet scoped, select:
- Inventory (if > $10M)
- Goodwill (if exists)
- Derivative Financial Instruments (all)

Result: 4 packs, specific FSLis only
Additional Coverage: 5.8%
Cumulative: 81.0%
```

### Phase 3: Complete Pack Inclusion

**Regulatory Requirements:**
```
All UK entities must be fully tested (regulation)

Action: Select all UK packs entirely
Result: 15 UK packs, all FSLis
Additional Coverage: 8.5%
Final Coverage: 89.5%
```

### Documentation

**Scoping Memo:**
```
Scope Methodology:
1. Automatic threshold scoping:
   - Net Revenue > $300M: 4 entities
   - Total Assets > $500M: 2 entities
   - Subtotal: 6 entities, 75.2% coverage

2. Risk-based manual selection:
   - Inventory > $10M: 3 entities
   - All derivatives: 2 entities (different from above)
   - Subtotal: 4 entities, 5.8% additional coverage

3. Regulatory requirement:
   - All UK entities: 15 entities
   - Subtotal: 8.5% additional coverage

Total Scope:
- Entities: 25 out of 48 (52%)
- Coverage: 89.5%
- Untested: 10.5%

Rationale for untested:
- Small immaterial entities
- Low risk profile
- No significant changes year-over-year
```

---

## Example 6: Complex Multi-Division Analysis

### Scenario
Large multinational with 10 divisions, need to ensure minimum coverage per division.

### Workbook Structure
```
Divisions: UK, US, Canada, France, Germany, Spain, Italy, Australia, Japan, Brazil
Total Entities: 150
Total FSLis: 250
```

### Division Coverage Requirements

**Requirement:** Each division must have minimum 70% coverage

### Power BI Analysis

#### Step 1: Division-Level Dashboard

**Visual Setup:**
- Slicer: Division (multi-select)
- Table: Coverage by Division
- Cards: Division-specific metrics

#### Step 2: Check Each Division

**UK Division:**
```
Entities: 25
Coverage (after threshold): 82%
Status: ✓ Meets requirement
```

**US Division:**
```
Entities: 40
Coverage (after threshold): 78%
Status: ✓ Meets requirement
```

**Spain Division:**
```
Entities: 8
Coverage (after threshold): 45%
Status: ✗ Below requirement
Action Needed: Manual selection required
```

#### Step 3: Address Deficiencies

**Spain Division - Manual Scoping:**
```
Select FSLis:
- Revenue: All Spain entities
- Accounts Receivable: Top 3 entities
- Inventory: Top 2 entities

Result: Coverage increases to 73%
Status: ✓ Now meets requirement
```

#### Step 4: Document by Division

**Division Coverage Summary:**
| Division | Entities | Threshold Scope | Manual Additions | Total Coverage | Status |
|----------|----------|-----------------|------------------|----------------|--------|
| UK | 25 | 82% | 0% | 82% | ✓ |
| US | 40 | 78% | 0% | 78% | ✓ |
| Canada | 12 | 71% | 0% | 71% | ✓ |
| France | 15 | 69% | 3% | 72% | ✓ |
| Germany | 18 | 74% | 0% | 74% | ✓ |
| Spain | 8 | 45% | 28% | 73% | ✓ |
| Italy | 10 | 68% | 4% | 72% | ✓ |
| Australia | 9 | 76% | 0% | 76% | ✓ |
| Japan | 7 | 81% | 0% | 81% | ✓ |
| Brazil | 6 | 52% | 20% | 72% | ✓ |

**Overall:**
```
Total Entities Scoped: 85 out of 150 (57%)
Average Coverage: 75.1%
All Divisions: ✓ Meet minimum 70% requirement
```

---

## Common Scenarios

### Scenario A: First-Time User

**Situation:** Never used the tool before

**Approach:**
1. Read Installation Guide
2. Create simple test workbook (3-5 tabs)
3. Run tool on test data
4. Review output tables
5. Import to Power BI (optional)
6. Then use on actual consolidation workbook

**Time Investment:** 2-3 hours including reading documentation

### Scenario B: Year-End Audit Scoping

**Situation:** Annual audit, need comprehensive scope

**Approach:**
1. Use hybrid scoping (threshold + manual + complete packs)
2. Set high thresholds for automatic inclusion
3. Manually add high-risk items
4. Include all new entities completely
5. Document thoroughly for audit file

**Coverage Target:** 80-90%

### Scenario C: Quarterly Review Scope

**Situation:** Q1, Q2, Q3 reviews, lighter testing

**Approach:**
1. Focus on threshold scoping only
2. Lower thresholds than year-end
3. Quick turnaround needed

**Coverage Target:** 60-70%

### Scenario D: New Entity Integration

**Situation:** Acquisition completed mid-year

**Approach:**
1. Run tool on updated consolidation
2. Identify new entities in Pack Number Company Table
3. Manually select ALL FSLis for new entities
4. Threshold scope existing entities
5. Document integration testing

### Scenario E: System Change Impact

**Situation:** New ERP implemented in specific division

**Approach:**
1. Filter to affected division
2. Select all entities in that division
3. Select all FSLis (complete testing)
4. Compare to prior period data
5. Document system change testing

### Scenario F: Discontin Operations Disposal

**Situation:** Discontinued operations sold during year

**Approach:**
1. Ensure Discontinued tab categorized correctly
2. Tool creates separate Discontinued Table
3. Review discontinued entities
4. Test discontinued FSLis through disposal date
5. Document disposal accounting

---

## Tips for Effective Usage

### Tip 1: Consistent Naming
- Use consistent workbook naming convention
- Example: `[Entity]_Consolidation_[Period].xlsx`
- Easier to track and document

### Tip 2: Save Categorization
- After first run, note your categorization
- Use same categories for subsequent periods
- Creates consistency

### Tip 3: Power BI Templates
- Create Power BI template with all measures
- Import new data each period
- Refresh and reuse scoping logic

### Tip 4: Document Methodology
- Create standard scoping memo template
- Include threshold rationale
- Document risk-based selections

### Tip 5: Validate Output
- Always verify a sample of data
- Check Pack Number Company Table for completeness
- Verify FSLi names match expectations

### Tip 6: Performance
- For large workbooks (100+ entities):
  - Close other applications
  - Run during off-peak hours
  - Consider processing segments separately

### Tip 7: Error Recovery
- Save consolidation workbook before running
- If tool errors, check categorization
- Restart Excel if needed
- Re-run from beginning

---

## Conclusion

These examples demonstrate the flexibility and power of the TGK Scoping Tool across various audit scenarios. The key to success is:

1. **Understanding your data structure**
2. **Proper tab categorization**
3. **Thoughtful scoping methodology**
4. **Thorough documentation**
5. **Regular validation**

Refer to [DOCUMENTATION.md](DOCUMENTATION.md) for technical details and [POWERBI_INTEGRATION_GUIDE.md](POWERBI_INTEGRATION_GUIDE.md) for Power BI specifics.

---

**Examples Version:** 1.0.0  
**Last Updated:** 2024  
**Tool Version:** 1.0.0
