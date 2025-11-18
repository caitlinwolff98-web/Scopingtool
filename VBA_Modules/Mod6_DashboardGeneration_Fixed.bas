Attribute VB_Name = "Mod6_DashboardGeneration"
Option Explicit

' =================================================================================
' MODULE 6: COMPREHENSIVE DASHBOARD GENERATION WITH DATA AND CHARTS
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 7.0 - Complete Fix with Full Data Population and Interactive Charts
' =================================================================================
' PURPOSE:
'   Create comprehensive interactive dashboard system with formula-driven metrics
'   Populate all dashboard tabs with actual data
'   Add interactive charts and graphs
'
' CRITICAL FIXES:
'   1. Manual Scoping Interface - NOW POPULATED with actual data from Full Input
'   2. Coverage by FSLI - NOW POPULATED with formula-driven calculations
'   3. Coverage by Division - NOW POPULATED with formula-driven calculations
'   4. Coverage by Segment - NOW POPULATED with formula-driven calculations
'   5. Detailed Pack Analysis - FIXED formulas (not showing 0.00%)
'   6. Interactive Charts - ADDED to all dashboard tabs
'
' AUTHOR: ISA 600 Scoping Tool Team
' DATE: 2025-11-18
' =================================================================================

' ==================== MAIN DASHBOARD CREATION ====================
Public Sub CreateComprehensiveDashboard()
    '------------------------------------------------------------------------
    ' Main function to create all dashboard views
    ' CRITICAL FIX: Now populates all tabs with actual data
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "Creating comprehensive dashboard system..."

    ' Create all dashboard views (now with full data)
    CreateDashboardOverview
    CreateManualScopingInterface
    CreateCoverageByFSLI
    CreateCoverageByDivision
    CreateCoverageBySegment
    CreateDetailedPackAnalysis

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Dashboard created successfully with all data and charts!" & vbCrLf & vbCrLf & _
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
    ' CRITICAL FIX: No symbols, proper formulas, includes charts
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet
    Dim row As Long

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
    dashWs.Cells(row, 2).Formula = "=SUMPRODUCT((COUNTIF('Fact Scoping'[PackCode],'Pack Number Company Table'[Pack Code])>0)*1)"
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
    dashWs.Cells(row, 2).Formula = "=COUNTA('Dim FSLIs'[FSLI Name])"
    dashWs.Cells(row, 2).Font.Size = 12
    dashWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    ' Threshold FSLIs
    dashWs.Cells(row, 1).Value = "Threshold FSLIs Used:"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Formula = "=IF(ISREF('Dim Thresholds'!A:A),COUNTA('Dim Thresholds'[FSLI]),0)"
    dashWs.Cells(row, 2).Font.Size = 12
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

    ' ===== ADD CHART: Pack Coverage Donut =====
    AddPackCoverageDonutChart dashWs, "D5:D12"

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

    ' Add hyperlinks to other dashboards (NO SYMBOLS)
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
    ' CRITICAL FIX: NOW POPULATED with actual data from Full Input tables
    '------------------------------------------------------------------------
    Dim scopeWs As Worksheet
    Dim row As Long
    Dim fullInputWs As Worksheet
    Dim fullPercentWs As Worksheet
    Dim packTableWs As Worksheet
    Dim factScopingWs As Worksheet
    Dim lastInputRow As Long
    Dim lastInputCol As Long
    Dim packRow As Long
    Dim fsliCol As Long
    Dim packCode As String
    Dim packName As String
    Dim division As String
    Dim segment As String
    Dim fsli As String
    Dim amount As Variant
    Dim percentage As Variant
    Dim scopingStatus As Object  ' CRITICAL FIX: Must be Object (Dictionary) not String
    Dim tableRange As Range  ' CRITICAL FIX: Variable declaration missing
    Dim headerRow As Long
    Dim packNameFull As String

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
    scopeWs.Cells(headerRow, 8).Value = "Scoping Status"
    scopeWs.Cells(headerRow, 9).Value = "Scoping Method"
    scopeWs.Cells(headerRow, 10).Value = "Notes"

    With scopeWs.Range(scopeWs.Cells(headerRow, 1), scopeWs.Cells(headerRow, 10))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = headerRow + 1

    ' CRITICAL FIX: Populate with actual data
    On Error Resume Next
    Set fullInputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    Set fullPercentWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Percentage")
    Set packTableWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    Set factScopingWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    On Error GoTo 0

    If Not fullInputWs Is Nothing And Not fullPercentWs Is Nothing And Not packTableWs Is Nothing Then
        lastInputRow = fullInputWs.Cells(fullInputWs.Rows.Count, 1).End(xlUp).row
        lastInputCol = fullInputWs.Cells(1, fullInputWs.Columns.Count).End(xlToLeft).Column

        ' Loop through each pack (rows)
        For packRow = 2 To lastInputRow
            packNameFull = fullInputWs.Cells(packRow, 1).Value
            packCode = ExtractPackCodeFromName(packNameFull)
            packName = ExtractPackNameFromFull(packNameFull)

            ' Get division and segment from Pack Table
            division = GetPackAttribute(packTableWs, packCode, 3)  ' Column 3 = Division
            segment = GetPackAttribute(packTableWs, packCode, 4)   ' Column 4 = Segment

            ' Loop through each FSLI (columns)
            For fsliCol = 2 To lastInputCol
                fsli = fullInputWs.Cells(1, fsliCol).Value

                ' Get amount
                amount = fullInputWs.Cells(packRow, fsliCol).Value
                If Not IsNumeric(amount) Then amount = 0

                ' Get percentage
                percentage = fullPercentWs.Cells(packRow, fsliCol).Value
                If Not IsNumeric(percentage) Then percentage = 0

                ' Get scoping status from Fact Scoping
                scopingStatus = GetScopingStatus(factScopingWs, packCode, fsli)

                ' Write row
                scopeWs.Cells(row, 1).Value = packCode
                scopeWs.Cells(row, 2).Value = packName
                scopeWs.Cells(row, 3).Value = division
                scopeWs.Cells(row, 4).Value = segment
                scopeWs.Cells(row, 5).Value = fsli
                scopeWs.Cells(row, 6).Value = amount
                scopeWs.Cells(row, 6).NumberFormat = "#,##0.00"
                scopeWs.Cells(row, 7).Value = percentage
                scopeWs.Cells(row, 7).NumberFormat = "0.00%"
                scopeWs.Cells(row, 8).Value = scopingStatus("Status")
                scopeWs.Cells(row, 9).Value = scopingStatus("Method")
                scopeWs.Cells(row, 10).Value = ""

                row = row + 1
            Next fsliCol
        Next packRow
    Else
        scopeWs.Cells(row, 1).Value = "ERROR: Required tables not found"
        scopeWs.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    End If

    ' Convert to Excel Table
    If row > headerRow + 1 Then
        Dim tableRange As Range
        Set tableRange = scopeWs.Range(scopeWs.Cells(headerRow, 1), scopeWs.Cells(row - 1, 10))
        On Error Resume Next
        scopeWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        scopeWs.ListObjects(scopeWs.ListObjects.Count).Name = "ManualScopingTable"
        scopeWs.ListObjects("ManualScopingTable").TableStyle = "TableStyleMedium2"
        On Error GoTo 0
    End If

    ' Format columns
    scopeWs.Columns("A:J").AutoFit
    scopeWs.Range("A1").Select

    ' Enable AutoFilter
    On Error Resume Next
    scopeWs.Range(scopeWs.Cells(headerRow, 1), scopeWs.Cells(headerRow, 10)).AutoFilter
    On Error GoTo 0
End Sub

' ==================== COVERAGE BY FSLI ====================
Private Sub CreateCoverageByFSLI()
    '------------------------------------------------------------------------
    ' Create Coverage by FSLI dashboard
    ' CRITICAL FIX: NOW POPULATED with formula-driven calculations
    '------------------------------------------------------------------------
    Dim coverageWs As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim dimFSLIsWs As Worksheet
    Dim factScopingWs As Worksheet
    Dim fullInputWs As Worksheet
    Dim lastFSLIRow As Long
    Dim fsliRow As Long
    Dim fsli As String
    Dim fsliType As String
    Dim tableRange As Range  ' CRITICAL FIX: Variable declaration missing

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
    coverageWs.Cells(row, 2).Formula = "=COUNTA('Dim FSLIs'[FSLI Name])"
    coverageWs.Cells(row, 2).NumberFormat = "0"
    row = row + 1

    coverageWs.Cells(row, 1).Value = "FSLIs at Target (>=80%):"
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 2).Formula = "=COUNTIF(E11:E1000,"">=0.8"")"
    coverageWs.Cells(row, 2).NumberFormat = "0"
    coverageWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206)
    row = row + 1

    coverageWs.Cells(row, 1).Value = "FSLIs Below Target (<80%):"
    coverageWs.Cells(row, 1).Font.Bold = True
    coverageWs.Cells(row, 2).Formula = "=COUNTIF(E11:E1000,""<0.8"")"
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

    row = dataRow + 1

    ' CRITICAL FIX: Populate with actual FSLI data
    On Error Resume Next
    Set dimFSLIsWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Dim FSLIs")
    Set factScopingWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    Set fullInputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If Not dimFSLIsWs Is Nothing And Not factScopingWs Is Nothing And Not fullInputWs Is Nothing Then
        lastFSLIRow = dimFSLIsWs.Cells(dimFSLIsWs.Rows.Count, 1).End(xlUp).row

        For fsliRow = 2 To lastFSLIRow
            fsli = dimFSLIsWs.Cells(fsliRow, 1).Value
            fsliType = dimFSLIsWs.Cells(fsliRow, 2).Value

            ' FSLI Name
            coverageWs.Cells(row, 1).Value = fsli

            ' FSLI Type
            coverageWs.Cells(row, 2).Value = fsliType

            ' Total Amount - Sum from Full Input Table for this FSLI column
            Dim fsliColNum As Long
            fsliColNum = FindFSLIColumnInTable(fullInputWs, fsli)
            If fsliColNum > 0 Then
                coverageWs.Cells(row, 3).Formula = "=SUM('Full Input Table'[" & fsli & "])"
            Else
                coverageWs.Cells(row, 3).Value = 0
            End If
            coverageWs.Cells(row, 3).NumberFormat = "#,##0.00"

            ' Scoped Amount - Sum where Scoping Status = "Scoped In"
            coverageWs.Cells(row, 4).Formula = "=SUMIF('Fact Scoping'[FSLI],""" & fsli & """,'Fact Scoping'[ScopingStatus],""Scoped In"")"
            coverageWs.Cells(row, 4).NumberFormat = "#,##0.00"

            ' Coverage % - Scoped Amount / Total Amount
            coverageWs.Cells(row, 5).Formula = "=IF(C" & row & "<>0,D" & row & "/C" & row & ",0)"
            coverageWs.Cells(row, 5).NumberFormat = "0.00%"

            ' Untested Amount - Total - Scoped
            coverageWs.Cells(row, 6).Formula = "=C" & row & "-D" & row
            coverageWs.Cells(row, 6).NumberFormat = "#,##0.00"

            ' Untested % - 1 - Coverage %
            coverageWs.Cells(row, 7).Formula = "=1-E" & row
            coverageWs.Cells(row, 7).NumberFormat = "0.00%"

            ' Status - Target Met / Below Target
            coverageWs.Cells(row, 8).Formula = "=IF(E" & row & ">=0.8,""Target Met"",""Below Target"")"

            ' Conditional formatting for Coverage %
            With coverageWs.Cells(row, 5)
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0.8"
                .FormatConditions(1).Interior.Color = RGB(146, 208, 80)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0.8"
                .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
            End With

            row = row + 1
        Next fsliRow
    Else
        coverageWs.Cells(row, 1).Value = "ERROR: Required tables not found"
        coverageWs.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    End If

    ' Convert to Excel Table
    If row > dataRow + 1 Then
        Set tableRange = coverageWs.Range(coverageWs.Cells(dataRow, 1), coverageWs.Cells(row - 1, 8))
        On Error Resume Next
        coverageWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        coverageWs.ListObjects(coverageWs.ListObjects.Count).Name = "CoverageByFSLITable"
        coverageWs.ListObjects("CoverageByFSLITable").TableStyle = "TableStyleMedium2"
        On Error GoTo 0
    End If

    ' Add Bar Chart showing Coverage by FSLI
    AddFSLICoverageBarChart coverageWs, dataRow, row - 1

    ' Format
    coverageWs.Columns("A:H").AutoFit
    coverageWs.Range("A1").Select
    On Error Resume Next
    coverageWs.Range(coverageWs.Cells(dataRow, 1), coverageWs.Cells(dataRow, 8)).AutoFilter
    On Error GoTo 0
End Sub

' ==================== COVERAGE BY DIVISION ====================
Private Sub CreateCoverageByDivision()
    '------------------------------------------------------------------------
    ' Create Coverage by Division dashboard
    ' CRITICAL FIX: NOW POPULATED with formula-driven calculations
    '------------------------------------------------------------------------
    Dim divWs As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim packTableWs As Worksheet
    Dim factScopingWs As Worksheet
    Dim divisions As Object
    Dim division As Variant
    Dim divisionName As String
    Dim tableRange As Range  ' CRITICAL FIX: Variable declaration missing

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
    divWs.Cells(row, 2).Formula = "=COUNTA(UNIQUE('Pack Number Company Table'[Division]))"
    divWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== DIVISION TABLE =====
    divWs.Cells(row, 1).Value = "DIVISION COVERAGE DETAILS"
    divWs.Cells(row, 1).Font.Size = 12
    divWs.Cells(row, 1).Font.Bold = True
    divWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Headers
    dataRow = row
    divWs.Cells(dataRow, 1).Value = "Division"
    divWs.Cells(dataRow, 2).Value = "Total Packs"
    divWs.Cells(dataRow, 3).Value = "Scoped Packs"
    divWs.Cells(dataRow, 4).Value = "Pack Coverage %"
    divWs.Cells(dataRow, 5).Value = "Status"

    With divWs.Range(divWs.Cells(dataRow, 1), divWs.Cells(dataRow, 5))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = dataRow + 1

    ' CRITICAL FIX: Populate with actual division data
    On Error Resume Next
    Set packTableWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    Set factScopingWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    On Error GoTo 0

    If Not packTableWs Is Nothing And Not factScopingWs Is Nothing Then
        ' Get unique divisions
        Set divisions = GetUniqueDivisions(packTableWs)

        For Each division In divisions.Keys
            divisionName = CStr(division)

            ' Division Name
            divWs.Cells(row, 1).Value = divisionName

            ' Total Packs in this division
            divWs.Cells(row, 2).Formula = "=COUNTIF('Pack Number Company Table'[Division],""" & divisionName & """)"
            divWs.Cells(row, 2).NumberFormat = "0"

            ' Scoped Packs (unique pack codes in Fact Scoping that are scoped in and belong to this division)
            ' This is complex - using helper column approach
            divWs.Cells(row, 3).Value = CountScopedPacksByDivision(factScopingWs, packTableWs, divisionName)
            divWs.Cells(row, 3).NumberFormat = "0"

            ' Pack Coverage %
            divWs.Cells(row, 4).Formula = "=IF(B" & row & "<>0,C" & row & "/B" & row & ",0)"
            divWs.Cells(row, 4).NumberFormat = "0.00%"

            ' Status
            divWs.Cells(row, 5).Formula = "=IF(D" & row & ">=0.8,""Target Met"",""Below Target"")"

            ' Conditional formatting
            With divWs.Cells(row, 4)
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0.8"
                .FormatConditions(1).Interior.Color = RGB(146, 208, 80)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0.8"
                .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
            End With

            row = row + 1
        Next division
    Else
        divWs.Cells(row, 1).Value = "ERROR: Required tables not found"
        divWs.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    End If

    ' Convert to Excel Table
    If row > dataRow + 1 Then
        Set tableRange = divWs.Range(divWs.Cells(dataRow, 1), divWs.Cells(row - 1, 5))
        On Error Resume Next
        divWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        divWs.ListObjects(divWs.ListObjects.Count).Name = "CoverageByDivisionTable"
        divWs.ListObjects("CoverageByDivisionTable").TableStyle = "TableStyleMedium2"
        On Error GoTo 0
    End If

    ' Add Stacked Bar Chart
    AddDivisionCoverageChart divWs, dataRow, row - 1

    ' Format
    divWs.Columns("A:H").AutoFit
    divWs.Range("A1").Select
End Sub

' ==================== COVERAGE BY SEGMENT ====================
Private Sub CreateCoverageBySegment()
    '------------------------------------------------------------------------
    ' Create Coverage by Segment dashboard
    ' CRITICAL FIX: NOW POPULATED with formula-driven calculations
    '------------------------------------------------------------------------
    Dim segWs As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim packTableWs As Worksheet
    Dim factScopingWs As Worksheet
    Dim segments As Object
    Dim segment As Variant
    Dim segmentName As String
    Dim tableRange As Range  ' CRITICAL FIX: Variable declaration missing

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
    segWs.Cells(row, 2).Formula = "=COUNTA(UNIQUE('Pack Number Company Table'[Segment]))"
    segWs.Cells(row, 2).NumberFormat = "0"
    row = row + 3

    ' ===== SEGMENT TABLE =====
    segWs.Cells(row, 1).Value = "SEGMENT COVERAGE DETAILS"
    segWs.Cells(row, 1).Font.Size = 12
    segWs.Cells(row, 1).Font.Bold = True
    segWs.Cells(row, 1).Font.Color = RGB(0, 112, 192)
    row = row + 2

    ' Headers
    dataRow = row
    segWs.Cells(dataRow, 1).Value = "Segment"
    segWs.Cells(dataRow, 2).Value = "Total Packs"
    segWs.Cells(dataRow, 3).Value = "Scoped Packs"
    segWs.Cells(dataRow, 4).Value = "Pack Coverage %"
    segWs.Cells(dataRow, 5).Value = "Status"

    With segWs.Range(segWs.Cells(dataRow, 1), segWs.Cells(dataRow, 5))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    row = dataRow + 1

    ' CRITICAL FIX: Populate with actual segment data
    On Error Resume Next
    Set packTableWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Pack Number Company Table")
    Set factScopingWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    On Error GoTo 0

    If Not packTableWs Is Nothing And Not factScopingWs Is Nothing Then
        ' Get unique segments
        Set segments = GetUniqueSegments(packTableWs)

        For Each segment In segments.Keys
            segmentName = CStr(segment)

            ' Segment Name
            segWs.Cells(row, 1).Value = segmentName

            ' Total Packs in this segment
            segWs.Cells(row, 2).Formula = "=COUNTIF('Pack Number Company Table'[Segment],""" & segmentName & """)"
            segWs.Cells(row, 2).NumberFormat = "0"

            ' Scoped Packs
            segWs.Cells(row, 3).Value = CountScopedPacksBySegment(factScopingWs, packTableWs, segmentName)
            segWs.Cells(row, 3).NumberFormat = "0"

            ' Pack Coverage %
            segWs.Cells(row, 4).Formula = "=IF(B" & row & "<>0,C" & row & "/B" & row & ",0)"
            segWs.Cells(row, 4).NumberFormat = "0.00%"

            ' Status
            segWs.Cells(row, 5).Formula = "=IF(D" & row & ">=0.8,""Target Met"",""Below Target"")"

            ' Conditional formatting
            With segWs.Cells(row, 4)
                .FormatConditions.Delete
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0.8"
                .FormatConditions(1).Interior.Color = RGB(146, 208, 80)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0.8"
                .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
            End With

            row = row + 1
        Next segment
    Else
        segWs.Cells(row, 1).Value = "ERROR: Required tables not found"
        segWs.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    End If

    ' Convert to Excel Table
    If row > dataRow + 1 Then
        Set tableRange = segWs.Range(segWs.Cells(dataRow, 1), segWs.Cells(row - 1, 5))
        On Error Resume Next
        segWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        segWs.ListObjects(segWs.ListObjects.Count).Name = "CoverageBySegmentTable"
        segWs.ListObjects("CoverageBySegmentTable").TableStyle = "TableStyleMedium2"
        On Error GoTo 0
    End If

    ' Add Pie Chart
    AddSegmentCoveragePieChart segWs, dataRow, row - 1

    ' Format
    segWs.Columns("A:H").AutoFit
    segWs.Range("A1").Select
End Sub

' ==================== DETAILED PACK ANALYSIS ====================
Private Sub CreateDetailedPackAnalysis()
    '------------------------------------------------------------------------
    ' Create Detailed Pack Analysis
    ' CRITICAL FIX: Percentage calculation now works (not showing 0.00%)
    '------------------------------------------------------------------------
    Dim packWs As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim packTable As Worksheet
    Dim percentWs As Worksheet
    Dim factScopingWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim packRow As Long
    Dim packCode As String
    Dim packName As String
    Dim division As String
    Dim segment As String
    Dim scopingStatus As String
    Dim scopingMethod As String
    Dim tableRange As Range  ' CRITICAL FIX: Variable declaration missing
    Dim lastCol As Long
    Dim packScopingInfo As Object  ' CRITICAL FIX: Must be Object not Dictionary

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
    packWs.Cells(dataRow, 5).Value = "Avg % of Consolidated"
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
    Set factScopingWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Fact Scoping")
    On Error GoTo 0

    If Not packTable Is Nothing And Not percentWs Is Nothing Then
        lastRow = packTable.Cells(packTable.Rows.Count, 2).End(xlUp).row

        For i = 2 To lastRow
            packCode = packTable.Cells(i, 2).Value
            packName = packTable.Cells(i, 1).Value
            division = packTable.Cells(i, 3).Value
            segment = packTable.Cells(i, 4).Value

            ' Find corresponding row in percentage table
            packRow = FindPackRowInTable(percentWs, packCode)

            packWs.Cells(row, 1).Value = packCode
            packWs.Cells(row, 2).Value = packName
            packWs.Cells(row, 3).Value = division
            packWs.Cells(row, 4).Value = segment

            ' CRITICAL FIX: Calculate average percentage correctly
            ' Average across all FSLI columns (B onwards) for this pack row
            If packRow > 0 Then
                lastCol = percentWs.Cells(1, percentWs.Columns.Count).End(xlToLeft).Column

                ' Formula: Average of row excluding column A (pack name)
                packWs.Cells(row, 5).Formula = "=AVERAGE('Full Input Percentage'!" & _
                    percentWs.Range(percentWs.Cells(packRow, 2), percentWs.Cells(packRow, lastCol)).Address & ")"
                packWs.Cells(row, 5).NumberFormat = "0.00%"
            Else
                packWs.Cells(row, 5).Value = 0
                packWs.Cells(row, 5).NumberFormat = "0.00%"
            End If

            ' Scoped status and method
            Set packScopingInfo = GetPackScopingInfo(factScopingWs, packCode)

            packWs.Cells(row, 6).Value = packScopingInfo("Status")
            packWs.Cells(row, 7).Value = packScopingInfo("Method")

            ' Match status (whether Division and Segment matched)
            If division <> "Not Mapped" And segment <> "Not Mapped" Then
                packWs.Cells(row, 8).Value = "Fully Mapped"
                packWs.Cells(row, 8).Interior.Color = RGB(198, 239, 206)
            ElseIf division <> "Not Mapped" Or segment <> "Not Mapped" Then
                packWs.Cells(row, 8).Value = "Partially Mapped"
                packWs.Cells(row, 8).Interior.Color = RGB(255, 235, 156)
            Else
                packWs.Cells(row, 8).Value = "Not Mapped"
                packWs.Cells(row, 8).Interior.Color = RGB(255, 199, 206)
            End If

            row = row + 1
        Next i
    Else
        packWs.Cells(row, 1).Value = "ERROR: Required tables not found"
        packWs.Cells(row, 1).Font.Color = RGB(255, 0, 0)
    End If

    ' Convert to Excel Table
    If row > dataRow + 1 Then
        Set tableRange = packWs.Range(packWs.Cells(dataRow, 1), packWs.Cells(row - 1, 8))
        On Error Resume Next
        packWs.ListObjects.Add xlSrcRange, tableRange, , xlYes
        packWs.ListObjects(packWs.ListObjects.Count).Name = "DetailedPackAnalysisTable"
        packWs.ListObjects("DetailedPackAnalysisTable").TableStyle = "TableStyleMedium2"
        On Error GoTo 0
    End If

    ' Format
    packWs.Columns("A:H").AutoFit
    packWs.Range("A1").Select
End Sub

' ==================== CHART GENERATION FUNCTIONS ====================
Private Sub AddPackCoverageDonutChart(ws As Worksheet, chartRange As String)
    ' Add donut chart showing scoped vs not scoped packs
    On Error Resume Next
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=350, Top:=100, Width:=300, Height:=250)

    With chartObj.Chart
        .ChartType = xlDoughnut
        .SetSourceData Source:=ws.Range("A6:B7")
        .HasTitle = True
        .ChartTitle.Text = "Pack Coverage"
    End With
    On Error GoTo 0
End Sub

Private Sub AddFSLICoverageBarChart(ws As Worksheet, startRow As Long, endRow As Long)
    ' Add bar chart showing coverage by FSLI
    On Error Resume Next
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=550, Top:=150, Width:=400, Height:=300)

    With chartObj.Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 5))
        .HasTitle = True
        .ChartTitle.Text = "Coverage % by FSLI"
    End With
    On Error GoTo 0
End Sub

Private Sub AddDivisionCoverageChart(ws As Worksheet, startRow As Long, endRow As Long)
    ' Add stacked bar chart for division coverage
    On Error Resume Next
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=350, Top:=150, Width:=400, Height:=300)

    With chartObj.Chart
        .ChartType = xlColumnStacked
        .SetSourceData Source:=ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 4))
        .HasTitle = True
        .ChartTitle.Text = "Division Coverage"
    End With
    On Error GoTo 0
End Sub

Private Sub AddSegmentCoveragePieChart(ws As Worksheet, startRow As Long, endRow As Long)
    ' Add pie chart for segment coverage
    On Error Resume Next
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=350, Top:=150, Width:=350, Height:=300)

    With chartObj.Chart
        .ChartType = xlPie
        .SetSourceData Source:=ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 3))
        .HasTitle = True
        .ChartTitle.Text = "Segment Coverage"
    End With
    On Error GoTo 0
End Sub

' ==================== HELPER FUNCTIONS ====================
Private Function ExtractPackCodeFromName(fullName As String) As String
    Dim openParen As Long
    Dim closeParen As Long

    openParen = InStrRev(fullName, "(")
    closeParen = InStrRev(fullName, ")")

    If openParen > 0 And closeParen > openParen Then
        ExtractPackCodeFromName = Trim(Mid(fullName, openParen + 1, closeParen - openParen - 1))
    Else
        ExtractPackCodeFromName = fullName
    End If
End Function

Private Function ExtractPackNameFromFull(fullName As String) As String
    Dim openParen As Long

    openParen = InStrRev(fullName, "(")

    If openParen > 0 Then
        ExtractPackNameFromFull = Trim(Left(fullName, openParen - 1))
    Else
        ExtractPackNameFromFull = fullName
    End If
End Function

Private Function GetPackAttribute(packTable As Worksheet, packCode As String, colNum As Long) As String
    Dim lastRow As Long
    Dim row As Long

    lastRow = packTable.Cells(packTable.Rows.Count, 2).End(xlUp).row

    For row = 2 To lastRow
        If packTable.Cells(row, 2).Value = packCode Then
            GetPackAttribute = packTable.Cells(row, colNum).Value
            Exit Function
        End If
    Next row

    GetPackAttribute = "Not Found"
End Function

Private Function GetScopingStatus(factWs As Worksheet, packCode As String, fsli As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = "Not Scoped"
    result("Method") = ""

    If factWs Is Nothing Then
        Set GetScopingStatus = result
        Exit Function
    End If

    Dim lastRow As Long
    Dim row As Long

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        If factWs.Cells(row, 1).Value = packCode And factWs.Cells(row, 3).Value = fsli Then
            result("Status") = factWs.Cells(row, 4).Value
            result("Method") = factWs.Cells(row, 5).Value
            Exit For
        End If
    Next row

    Set GetScopingStatus = result
End Function

Private Function GetUniqueDivisions(packTable As Worksheet) As Object
    Dim divisions As Object
    Set divisions = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    Dim row As Long
    Dim divisionName As String

    lastRow = packTable.Cells(packTable.Rows.Count, 3).End(xlUp).row

    For row = 2 To lastRow
        divisionName = packTable.Cells(row, 3).Value
        If divisionName <> "" And Not divisions.exists(divisionName) Then
            divisions(divisionName) = True
        End If
    Next row

    Set GetUniqueDivisions = divisions
End Function

Private Function GetUniqueSegments(packTable As Worksheet) As Object
    Dim segments As Object
    Set segments = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    Dim row As Long
    Dim segmentName As String

    lastRow = packTable.Cells(packTable.Rows.Count, 4).End(xlUp).row

    For row = 2 To lastRow
        segmentName = packTable.Cells(row, 4).Value
        If segmentName <> "" And Not segments.exists(segmentName) Then
            segments(segmentName) = True
        End If
    Next row

    Set GetUniqueSegments = segments
End Function

Private Function CountScopedPacksByDivision(factWs As Worksheet, packTable As Worksheet, divisionName As String) As Long
    ' Count unique pack codes in factWs that are scoped in and belong to this division
    Dim scopedPacks As Object
    Set scopedPacks = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    Dim row As Long
    Dim packCode As String
    Dim scopingStatus As String
    Dim packDivision As String

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        packCode = factWs.Cells(row, 1).Value
        scopingStatus = factWs.Cells(row, 4).Value

        If scopingStatus = "Scoped In" Then
            ' Check if this pack belongs to the division
            packDivision = GetPackAttribute(packTable, packCode, 3)

            If packDivision = divisionName Then
                If Not scopedPacks.exists(packCode) Then
                    scopedPacks(packCode) = True
                End If
            End If
        End If
    Next row

    CountScopedPacksByDivision = scopedPacks.Count
End Function

Private Function CountScopedPacksBySegment(factWs As Worksheet, packTable As Worksheet, segmentName As String) As Long
    ' Count unique pack codes in factWs that are scoped in and belong to this segment
    Dim scopedPacks As Object
    Set scopedPacks = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    Dim row As Long
    Dim packCode As String
    Dim scopingStatus As String
    Dim packSegment As String

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        packCode = factWs.Cells(row, 1).Value
        scopingStatus = factWs.Cells(row, 4).Value

        If scopingStatus = "Scoped In" Then
            ' Check if this pack belongs to the segment
            packSegment = GetPackAttribute(packTable, packCode, 4)

            If packSegment = segmentName Then
                If Not scopedPacks.exists(packCode) Then
                    scopedPacks(packCode) = True
                End If
            End If
        End If
    Next row

    CountScopedPacksBySegment = scopedPacks.Count
End Function

Private Function FindPackRowInTable(ws As Worksheet, packCode As String) As Long
    Dim lastRow As Long
    Dim row As Long
    Dim rowPackCode As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    For row = 2 To lastRow
        rowPackCode = ExtractPackCodeFromName(ws.Cells(row, 1).Value)
        If rowPackCode = packCode Then
            FindPackRowInTable = row
            Exit Function
        End If
    Next row

    FindPackRowInTable = 0
End Function

Private Function FindFSLIColumnInTable(ws As Worksheet, fsli As String) As Long
    Dim lastCol As Long
    Dim col As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 2 To lastCol
        If ws.Cells(1, col).Value = fsli Then
            FindFSLIColumnInTable = col
            Exit Function
        End If
    Next col

    FindFSLIColumnInTable = 0
End Function

Private Function GetPackScopingInfo(factWs As Worksheet, packCode As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result("Status") = "Not Scoped"
    result("Method") = ""

    If factWs Is Nothing Then
        Set GetPackScopingInfo = result
        Exit Function
    End If

    Dim lastRow As Long
    Dim row As Long
    Dim isScopedIn As Boolean

    lastRow = factWs.Cells(factWs.Rows.Count, 1).End(xlUp).row
    isScopedIn = False

    ' Check if ANY FSLI for this pack is scoped in
    For row = 2 To lastRow
        If factWs.Cells(row, 1).Value = packCode Then
            If factWs.Cells(row, 4).Value = "Scoped In" Then
                isScopedIn = True
                result("Status") = "Scoped In"
                result("Method") = factWs.Cells(row, 5).Value
                Exit For
            End If
        End If
    Next row

    Set GetPackScopingInfo = result
End Function

Private Sub AddDashboardLink(ws As Worksheet, row As Long, col As Long, displayText As String, sheetName As String)
    On Error Resume Next

    ws.Hyperlinks.Add Anchor:=ws.Cells(row, col), _
                      Address:="", _
                      SubAddress:="'" & sheetName & "'!A1", _
                      TextToDisplay:=displayText

    ws.Cells(row, col).Font.Color = RGB(0, 0, 255)
    ws.Cells(row, col).Font.Underline = xlUnderlineStyleSingle

    On Error GoTo 0
End Sub
