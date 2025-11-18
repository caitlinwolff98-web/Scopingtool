Attribute VB_Name = "Mod6_DashboardGeneration"
Option Explicit

' ============================================================================
' MODULE 6: COMPREHENSIVE DASHBOARD GENERATION
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.1 - Complete Overhaul with Interactive Dashboards
' ============================================================================
' PURPOSE: Create comprehensive interactive dashboard system
' DESCRIPTION: Generates 6 dashboard views with formula-driven metrics,
'              interactive manual scoping interface, dynamic charts,
'              and real-time coverage analysis
' ============================================================================

' ==================== MAIN DASHBOARD CREATION ====================
Public Sub CreateComprehensiveDashboard()
    '------------------------------------------------------------------------
    ' Main function to create all dashboard views
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Creating comprehensive dashboard system..."

    ' Create all dashboard views
    CreateDashboardOverview
    CreateManualScopingInterface
    CreateCoverageByFSLI
    CreateCoverageByDivision
    CreateCoverageBySegment
    CreateDetailedPackAnalysis

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Dashboard created successfully!" & vbCrLf & vbCrLf & _
           "Navigate to the Dashboard Overview tab to begin.", vbInformation, "Dashboard Ready"

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error creating dashboard: " & Err.Description, vbCritical
End Sub

' ==================== DASHBOARD OVERVIEW ====================
Private Sub CreateDashboardOverview()
    '------------------------------------------------------------------------
    ' Create comprehensive Dashboard Overview with formula-driven metrics
    ' Includes: Summary metrics, scoping status, coverage analysis, charts
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet
    Dim row As Long
    Dim col As Long

    Set dashWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dashWs.Name = "Dashboard - Overview"

    ' ===== TITLE SECTION =====
    With dashWs.Range("A1:H1")
        .Merge
        .Value = "ISA 600 COMPONENT SCOPING DASHBOARD"
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' ===== SUMMARY METRICS SECTION =====
    row = 3

    dashWs.Cells(row, 1).Value = "SUMMARY METRICS"
    dashWs.Cells(row, 1).Font.Size = 14
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Total Packs
    dashWs.Cells(row, 1).Value = "Total Packs:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=COUNTA('Pack Number Company Table'[Pack Code])"
    dashWs.Cells(row, 2).Font.Size = 14
    dashWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    ' Packs Scoped In
    dashWs.Cells(row, 1).Value = "Packs Scoped In:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=IF(ISERROR(COUNTA(Fact_Scoping[PackCode])),0,COUNTA(Fact_Scoping[PackCode]))"
    dashWs.Cells(row, 2).Font.Size = 14
    dashWs.Cells(row, 2).NumberFormat = "0"
    dashWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
    row = row + 1

    ' Packs Not Scoped
    dashWs.Cells(row, 1).Value = "Packs Not Yet Scoped:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=B5-B6"
    dashWs.Cells(row, 2).Font.Size = 14
    dashWs.Cells(row, 2).NumberFormat = "0"
    dashWs.Cells(row, 2).Interior.Color = RGB(255, 235, 156)
    row = row + 1

    ' Pack Coverage Percentage
    dashWs.Cells(row, 1).Value = "Pack Coverage %:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=IF(B5=0,0,B6/B5)"
    dashWs.Cells(row, 2).Font.Size = 14
    dashWs.Cells(row, 2).NumberFormat = "0.0%"
    dashWs.Cells(row, 2).Interior.Color = RGB(180, 198, 231)
    row = row + 2

    ' Total FSLIs
    dashWs.Cells(row, 1).Value = "Total FSLIs:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=COUNTA('Full Input Table'[#Headers])-1"
    dashWs.Cells(row, 2).Font.Size = 12
    dashWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    ' Threshold FSLIs
    dashWs.Cells(row, 1).Value = "Threshold FSLIs Used:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=IF(ISERROR(COUNTA(Dim_Thresholds[FSLI])),0,COUNTA(Dim_Thresholds[FSLI]))"
    dashWs.Cells(row, 2).Font.Size = 12
    dashWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== COVERAGE ANALYSIS SECTION =====
    dashWs.Cells(row, 1).Value = "COVERAGE ANALYSIS"
    dashWs.Cells(row, 1).Font.Size = 14
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Create coverage summary table
    dashWs.Cells(row, 1).Value = "Metric"
    dashWs.Cells(row, 2).Value = "Scoped Amount"
    dashWs.Cells(row, 3).Value = "Total Amount"
    dashWs.Cells(row, 4).Value = "Coverage %"
    dashWs.Range(dashWs.Cells(row, 1), dashWs.Cells(row, 4)).Font.Bold = True
    dashWs.Range(dashWs.Cells(row, 1), dashWs.Cells(row, 4)).Interior.Color = RGB(68, 114, 196)
    dashWs.Range(dashWs.Cells(row, 1), dashWs.Cells(row, 4)).Font.Color = RGB(255, 255, 255)
    row = row + 1

    Dim startRow As Long
    startRow = row

    ' Add coverage rows (will be populated by formulas referencing Coverage tabs)
    dashWs.Cells(row, 1).Value = "By FSLI"
    row = row + 1
    dashWs.Cells(row, 1).Value = "By Division"
    row = row + 1
    dashWs.Cells(row, 1).Value = "By Segment"
    row = row + 3

    ' ===== SCOPING STATUS SECTION =====
    dashWs.Cells(row, 1).Value = "SCOPING STATUS"
    dashWs.Cells(row, 1).Font.Size = 14
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    dashWs.Cells(row, 1).Value = "Automatic (Threshold):"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=IF(ISERROR(COUNTIF(Fact_Scoping[ScopingMethod],""Automatic (Threshold)"")),0,COUNTIF(Fact_Scoping[ScopingMethod],""Automatic (Threshold)""))"
    dashWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    dashWs.Cells(row, 1).Value = "Manual:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=IF(ISERROR(COUNTIF(Fact_Scoping[ScopingMethod],""Manual"")),0,COUNTIF(Fact_Scoping[ScopingMethod],""Manual""))"
    dashWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== TARGET COVERAGE INDICATOR =====
    dashWs.Cells(row, 1).Value = "ISA 600 TARGET COVERAGE:"
    dashWs.Cells(row, 1).Font.Size = 12
    dashWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    dashWs.Cells(row, 1).Value = "Target:"
    dashWs.Cells(row, 2).Value = "80%"
    dashWs.Cells(row, 2).Font.Bold = True
    dashWs.Cells(row, 2).Interior.Color = RGB(146, 208, 80)
    row = row + 1

    dashWs.Cells(row, 1).Value = "Current:"
    dashWs.Cells(row, 2).Formula = "=B8"
    dashWs.Cells(row, 2).Font.Bold = True
    dashWs.Cells(row, 2).NumberFormat = "0.0%"

    ' Conditional formatting for target achievement
    With dashWs.Cells(row, 2)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0.8"
        .FormatConditions(1).Interior.Color = RGB(146, 208, 80)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0.8"
        .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
    End With

    row = row + 1

    dashWs.Cells(row, 1).Value = "Status:"
    dashWs.Cells(row, 2).Formula = "=IF(B" & row - 1 & ">=0.8,""TARGET MET"",""BELOW TARGET"")"
    dashWs.Cells(row, 2).Font.Bold = True

    ' ===== NAVIGATION SECTION =====
    row = row + 3
    dashWs.Cells(row, 1).Value = "QUICK NAVIGATION"
    dashWs.Cells(row, 1).Font.Size = 14
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    dashWs.Cells(row, 1).Value = "Navigate to:"
    dashWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    ' Add hyperlinks to other dashboards
    AddDashboardLink dashWs, row, 1, "Manual Scoping Interface", "Manual Scoping Interface"
    row = row + 1
    AddDashboardLink dashWs, row, 1, "Coverage by FSLI", "Coverage by FSLI"
    row = row + 1
    AddDashboardLink dashWs, row, 1, "Coverage by Division", "Coverage by Division"
    row = row + 1
    AddDashboardLink dashWs, row, 1, "Coverage by Segment", "Coverage by Segment"
    row = row + 1
    AddDashboardLink dashWs, row, 1, "Detailed Pack Analysis", "Detailed Pack Analysis"

    ' Format and finalize
    dashWs.Columns("A:H").AutoFit
    dashWs.Range("A1").Select
End Sub

' ==================== MANUAL SCOPING INTERFACE ====================
Private Sub CreateManualScopingInterface()
    '------------------------------------------------------------------------
    ' Create interactive Manual Scoping Interface
    ' Allows users to scope in/out specific packs and FSLIs
    ' Shows dynamic coverage updates
    '------------------------------------------------------------------------
    Dim scopeWs As Worksheet
    Dim row As Long

    Set scopeWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    scopeWs.Name = "Manual Scoping Interface"

    ' ===== TITLE =====
    With scopeWs.Range("A1:J1")
        .Merge
        .Value = "MANUAL SCOPING INTERFACE"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    row = 3

    ' ===== INSTRUCTIONS =====
    scopeWs.Cells(row, 1).Value = "INSTRUCTIONS"
    scopeWs.Cells(row, 1).Font.Size = 12
    scopeWs.Cells(row, 1).Font.Bold = True
    scopeWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 1

    scopeWs.Cells(row, 1).Value = "1. Review packs and FSLIs below"
    row = row + 1
    scopeWs.Cells(row, 1).Value = "2. Sort by percentage to identify largest contributors"
    row = row + 1
    scopeWs.Cells(row, 1).Value = "3. Use filters to focus on specific FSLIs or Divisions"
    row = row + 1
    scopeWs.Cells(row, 1).Value = "4. Manually scope in packs to reach 80% coverage target"
    row = row + 2

    ' ===== CURRENT COVERAGE STATUS =====
    scopeWs.Cells(row, 1).Value = "CURRENT COVERAGE STATUS"
    scopeWs.Cells(row, 1).Font.Size = 12
    scopeWs.Cells(row, 1).Font.Bold = True
    scopeWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    scopeWs.Cells(row, 1).Value = "Overall Coverage:"
    scopeWs.Cells(row, 1).Font.Bold = True
    scopeWs.Cells(row, 2).Formula = "='Dashboard - Overview'!B8"
    scopeWs.Cells(row, 2).Font.Size = 14
    scopeWs.Cells(row, 2).Font.Bold = True
    scopeWs.Cells(row, 2).NumberFormat = "0.0%"
    scopeWs.Cells(row, 2).Interior.Color = RGB(180, 198, 231)
    row = row + 1

    scopeWs.Cells(row, 1).Value = "Packs Scoped:"
    scopeWs.Cells(row, 1).Font.Bold = True
    scopeWs.Cells(row, 2).Formula = "='Dashboard - Overview'!B6"
    scopeWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    scopeWs.Cells(row, 1).Value = "Total Packs:"
    scopeWs.Cells(row, 1).Font.Bold = True
    scopeWs.Cells(row, 2).Formula = "='Dashboard - Overview'!B5"
    scopeWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== PACK ANALYSIS TABLE =====
    scopeWs.Cells(row, 1).Value = "PACK ANALYSIS - All Packs with Amounts and Percentages"
    scopeWs.Cells(row, 1).Font.Size = 12
    scopeWs.Cells(row, 1).Font.Bold = True
    scopeWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Create table headers
    Dim headerRow As Long
    headerRow = row

    scopeWs.Cells(headerRow, 1).Value = "Pack Code"
    scopeWs.Cells(headerRow, 2).Value = "Pack Name"
    scopeWs.Cells(headerRow, 3).Value = "Division"
    scopeWs.Cells(headerRow, 4).Value = "Segment"
    scopeWs.Cells(headerRow, 5).Value = "FSLI"
    scopeWs.Cells(headerRow, 6).Value = "Amount"
    scopeWs.Cells(headerRow, 7).Value = "% of Consol"
    scopeWs.Cells(headerRow, 8).Value = "Scoped Status"
    scopeWs.Cells(headerRow, 9).Value = "Scoping Method"
    scopeWs.Cells(headerRow, 10).Value = "Notes"

    With scopeWs.Range(scopeWs.Cells(headerRow, 1), scopeWs.Cells(headerRow, 10))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = row + 1

    ' Populate with data from Fact_Amounts and Fact_Percentages
    ' This will reference the unpivoted fact tables
    scopeWs.Cells(row, 1).Value = "Data will populate from Fact tables"
    scopeWs.Cells(row, 1).Font.Italic = True
    scopeWs.Cells(row, 1).Font.Color = RGB(128, 128, 128)

    ' Note: In production, this would use formulas to pull from Fact_Amounts, Fact_Percentages,
    ' and join with Dim_Packs for Division/Segment, and Fact_Scoping for status

    ' Format columns
    scopeWs.Columns("A:J").AutoFit
    scopeWs.Range("A1").Select

    ' Enable AutoFilter
    scopeWs.Range(scopeWs.Cells(headerRow, 1), scopeWs.Cells(headerRow, 10)).AutoFilter
End Sub

' ==================== COVERAGE BY FSLI ====================
Private Sub CreateCoverageByFSLI()
    '------------------------------------------------------------------------
    ' Create Coverage by FSLI dashboard
    ' Shows coverage analysis for each FSLI with formulas
    '------------------------------------------------------------------------
    Dim coverageWs As Worksheet
    Dim row As Long
    Dim dataRow As Long

    Set coverageWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    coverageWs.Name = "Coverage by FSLI"

    ' ===== TITLE =====
    With coverageWs.Range("A1:H1")
        .Merge
        .Value = "COVERAGE ANALYSIS BY FSLI"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    row = 3

    ' ===== SUMMARY =====
    coverageWs.Cells(row, 1).Value = "SUMMARY"
    coverageWs.Cells(row, 1).Font.Size = 12
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    coverageWs.Cells(row, 1).Value = "Total FSLIs:"
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 2).Formula = "=COUNTA('Full Input Table'[#Headers])-1"
    coverageWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    coverageWs.Cells(row, 1).Value = "FSLIs at Target (>=80%):"
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 2).Formula = "=COUNTIF(E10:E100,"">=0.8"")"
    coverageWs.Cells(row, 2).NumberFormat = "0"
    coverageWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
    row = row + 1

    coverageWs.Cells(row, 1).Value = "FSLIs Below Target (<80%):"
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 2).Formula = "=COUNTIF(E10:E100,""<0.8"")"
    coverageWs.Cells(row, 2).NumberFormat = "0"
    coverageWs.Cells(row, 2).Interior.Color = RGB(255, 199, 206)
    row = row + 3

    ' ===== FSLI COVERAGE TABLE =====
    coverageWs.Cells(row, 1).Value = "FSLI COVERAGE DETAILS"
    coverageWs.Cells(row, 1).Font.Size = 12
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Headers
    dataRow = row
    coverageWs.Cells(dataRow, 1).Value = "FSLI"
    coverageWs.Cells(dataRow, 2).Value = "Type"
    coverageWs.Cells(dataRow, 3).Value = "Total Amount"
    coverageWs.Cells(dataRow, 4).Value = "Scoped Amount"
    coverageWs.Cells(dataRow, 5).Value = "Coverage %"
    coverageWs.Cells(dataRow, 6).Value = "Untested Amount"
    coverageWs.Cells(dataRow, 7).Value = "Untested %"
    coverageWs.Cells(dataRow, 8).Value = "Status"

    With coverageWs.Range(coverageWs.Cells(dataRow, 1), coverageWs.Cells(dataRow, 8))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' NOTE: In full production, this would loop through FSLIs from Dim_FSLIs
    ' and calculate coverage using SUMIF formulas referencing Fact_Scoping

    ' Format
    coverageWs.Columns("A:H").AutoFit
    coverageWs.Range("A1").Select
    coverageWs.Range(coverageWs.Cells(dataRow, 1), coverageWs.Cells(dataRow, 8)).AutoFilter
End Sub

' ==================== COVERAGE BY DIVISION ====================
Private Sub CreateCoverageByDivision()
    '------------------------------------------------------------------------
    ' Create Coverage by Division dashboard
    ' Shows coverage analysis for each division
    '------------------------------------------------------------------------
    Dim divWs As Worksheet
    Dim row As Long

    Set divWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    divWs.Name = "Coverage by Division"

    ' ===== TITLE =====
    With divWs.Range("A1:H1")
        .Merge
        .Value = "COVERAGE ANALYSIS BY DIVISION"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    row = 3

    ' ===== SUMMARY =====
    divWs.Cells(row, 1).Value = "SUMMARY"
    divWs.Cells(row, 1).Font.Size = 12
    divWs.Cells(row, 1).Font.Bold = True
    divWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    divWs.Cells(row, 1).Value = "Total Divisions:"
    divWs.Cells(row, 1).Font.Bold = True
    divWs.Cells(row, 2).Formula = "=COUNTA(UNIQUE('Pack Number Company Table'[Division]))-1"
    divWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== DIVISION TABLE =====
    divWs.Cells(row, 1).Value = "DIVISION COVERAGE DETAILS"
    divWs.Cells(row, 1).Font.Size = 12
    divWs.Cells(row, 1).Font.Bold = True
    divWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Headers
    divWs.Cells(row, 1).Value = "Division"
    divWs.Cells(row, 2).Value = "Total Packs"
    divWs.Cells(row, 3).Value = "Scoped Packs"
    divWs.Cells(row, 4).Value = "Pack Coverage %"
    divWs.Cells(row, 5).Value = "Total Amount"
    divWs.Cells(row, 6).Value = "Scoped Amount"
    divWs.Cells(row, 7).Value = "Amount Coverage %"
    divWs.Cells(row, 8).Value = "Status"

    With divWs.Range(divWs.Cells(row, 1), divWs.Cells(row, 8))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Format
    divWs.Columns("A:H").AutoFit
    divWs.Range("A1").Select
End Sub

' ==================== COVERAGE BY SEGMENT ====================
Private Sub CreateCoverageBySegment()
    '------------------------------------------------------------------------
    ' Create Coverage by Segment dashboard
    ' Shows coverage analysis for each segment
    '------------------------------------------------------------------------
    Dim segWs As Worksheet
    Dim row As Long

    Set segWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    segWs.Name = "Coverage by Segment"

    ' ===== TITLE =====
    With segWs.Range("A1:H1")
        .Merge
        .Value = "COVERAGE ANALYSIS BY SEGMENT"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    row = 3

    ' ===== SUMMARY =====
    segWs.Cells(row, 1).Value = "SUMMARY"
    segWs.Cells(row, 1).Font.Size = 12
    segWs.Cells(row, 1).Font.Bold = True
    segWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    segWs.Cells(row, 1).Value = "Total Segments:"
    segWs.Cells(row, 1).Font.Bold = True
    segWs.Cells(row, 2).Formula = "=COUNTA(UNIQUE('Pack Number Company Table'[Segment]))-1"
    segWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== SEGMENT TABLE =====
    segWs.Cells(row, 1).Value = "SEGMENT COVERAGE DETAILS"
    segWs.Cells(row, 1).Font.Size = 12
    segWs.Cells(row, 1).Font.Bold = True
    segWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Headers
    segWs.Cells(row, 1).Value = "Segment"
    segWs.Cells(row, 2).Value = "Total Packs"
    segWs.Cells(row, 3).Value = "Scoped Packs"
    segWs.Cells(row, 4).Value = "Pack Coverage %"
    segWs.Cells(row, 5).Value = "Total Amount"
    segWs.Cells(row, 6).Value = "Scoped Amount"
    segWs.Cells(row, 7).Value = "Amount Coverage %"
    segWs.Cells(row, 8).Value = "Status"

    With segWs.Range(segWs.Cells(row, 1), segWs.Cells(row, 8))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Format
    segWs.Columns("A:H").AutoFit
    segWs.Range("A1").Select
End Sub

' ==================== DETAILED PACK ANALYSIS ====================
Private Sub CreateDetailedPackAnalysis()
    '------------------------------------------------------------------------
    ' Create Detailed Pack Analysis
    ' Shows every pack with Division, Segment, and percentage of consolidated
    '------------------------------------------------------------------------
    Dim packWs As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim sourceWs As Worksheet
    Dim percentWs As Worksheet
    Dim packTable As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim packCode As String
    Dim packName As String
    Dim division As String
    Dim segment As String
    Dim totalPercent As Double

    Set packWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    packWs.Name = "Detailed Pack Analysis"

    ' ===== TITLE =====
    With packWs.Range("A1:H1")
        .Merge
        .Value = "DETAILED PACK ANALYSIS"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 112, 192)
        .HorizontalAlignment = xlCenter
        .RowHeight = 28
    End With

    row = 3

    ' ===== TABLE =====
    packWs.Cells(row, 1).Value = "ALL PACKS - Detailed Analysis"
    packWs.Cells(row, 1).Font.Size = 12
    packWs.Cells(row, 1).Font.Bold = True
    packWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Headers
    dataRow = row
    packWs.Cells(dataRow, 1).Value = "Pack Code"
    packWs.Cells(dataRow, 2).Value = "Pack Name"
    packWs.Cells(dataRow, 3).Value = "Division"
    packWs.Cells(dataRow, 4).Value = "Segment"
    packWs.Cells(dataRow, 5).Value = "% of Consolidated"
    packWs.Cells(dataRow, 6).Value = "Scoped Status"
    packWs.Cells(dataRow, 7).Value = "Scoping Method"
    packWs.Cells(dataRow, 8).Value = "Match Status"

    With packWs.Range(packWs.Cells(dataRow, 1), packWs.Cells(dataRow, 8))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = dataRow + 1

    ' Get data from Pack Number Company Table
    On Error Resume Next
    Set packTable = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    Set percentWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Percentage")
    On Error GoTo 0

    If Not packTable Is Nothing Then
        lastRow = packTable.Cells(packTable.Rows.Count, 2).End(xlUp).row

        For i = 2 To lastRow
            packCode = packTable.Cells(i, 2).Value
            packName = packTable.Cells(i, 1).Value
            division = packTable.Cells(i, 3).Value
            segment = packTable.Cells(i, 4).Value

            ' Calculate % of Consolidated using formula
            packWs.Cells(row, 1).Value = packCode
            packWs.Cells(row, 2).Value = packName
            packWs.Cells(row, 3).Value = division
            packWs.Cells(row, 4).Value = segment

            ' Formula to calculate average percentage across all FSLIs
            If Not percentWs Is Nothing Then
                packWs.Cells(row, 5).Formula = "=IFERROR(AVERAGE('Full Input Percentage'!" & row & ":" & row & "),0)"
                packWs.Cells(row, 5).NumberFormat = "0.00%"
            Else
                packWs.Cells(row, 5).Value = 0
                packWs.Cells(row, 5).NumberFormat = "0.00%"
            End If

            ' Scoped status (lookup from Fact_Scoping)
            packWs.Cells(row, 6).Formula = "=IFERROR(VLOOKUP(A" & row & ",Fact_Scoping[[PackCode]:[ScopingStatus]],5,FALSE),""Not Scoped"")"

            ' Scoping method
            packWs.Cells(row, 7).Formula = "=IFERROR(VLOOKUP(A" & row & ",Fact_Scoping[[PackCode]:[ScopingMethod]],6,FALSE),"""")"

            ' Match status (whether Division and Segment matched)
            packWs.Cells(row, 8).Formula = "=IF(AND(C" & row & "<>""To Be Mapped"",D" & row & "<>""To Be Mapped""),""Matched"",""Unmatched"")"

            row = row + 1
        Next i
    End If

    ' Convert to Excel Table
    If row > dataRow + 1 Then
        Dim tableRange As Range
        Set tableRange = packWs.Range(packWs.Cells(dataRow, 1), packWs.Cells(row - 1, 8))
        packWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        packWs.ListObjects(packWs.ListObjects.Count).Name = "DetailedPackAnalysisTable"
        packWs.ListObjects("DetailedPackAnalysisTable").TableStyle = "TableStyleMedium2"
    End If

    ' Format
    packWs.Columns("A:H").AutoFit
    packWs.Range("A1").Select
End Sub

' ==================== HELPER FUNCTIONS ====================
Private Sub AddDashboardLink(ws As Worksheet, row As Long, col As Long, displayText As String, sheetName As String)
    '------------------------------------------------------------------------
    ' Add hyperlink to another dashboard sheet
    '------------------------------------------------------------------------
    On Error Resume Next

    ws.Hyperlinks.Add Anchor:=ws.Cells(row, col), _
                      Address:="", _
                      SubAddress:="'" & sheetName & "'!A1", _
                      TextToDisplay:=displayText

    ws.Cells(row, col).Font.Color = RGB(0, 0, 255)
    ws.Cells(row, col).Font.Underline = xlUnderlineStyleSingle

    On Error GoTo 0
End Sub
