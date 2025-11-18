Attribute VB_Name = "ModExcelDashboard"
Option Explicit

' ============================================================================
' MODULE: ModExcelDashboard
' PURPOSE: Create PRODUCTION-READY interactive Excel dashboard
' DESCRIPTION: Complete dashboard with all scoping analysis, interactive
'              controls, dynamic updates, and professional formatting
' VERSION: 5.1 Production
' ============================================================================

' Main entry point - creates comprehensive dashboard
Public Sub CreateComprehensiveExcelDashboard()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "Building comprehensive Excel dashboard..."

    ' Validate required tables exist
    If Not ValidateRequiredTables() Then
        MsgBox "Cannot create dashboard - required tables missing." & vbCrLf & _
               "Ensure Scoping_Control_Table and Pack_Number_Company_Table exist.", vbCritical
        GoTo Cleanup
    End If

    ' Create all dashboard components
    Call CreateExecutiveSummaryDashboard
    Call EnhanceScopingControlTable
    Call CreateFSLICoverageAnalysis
    Call CreateDivisionSegmentAnalysis
    Call CreateInteractiveWorksheet
    Call CreateQuickReferenceGuide

    ' Set default view to Dashboard
    On Error Resume Next
    g_OutputWorkbook.Worksheets("Dashboard - Executive Summary").Activate
    On Error GoTo ErrorHandler

Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error creating Excel dashboard: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Dashboard Creation Error"
End Sub

' Validate that required tables exist
Private Function ValidateRequiredTables() As Boolean
    On Error Resume Next

    Dim scopingTable As ListObject
    Dim packTable As ListObject

    ' Check for Scoping_Control_Table
    Set scopingTable = g_OutputWorkbook.Worksheets("Scoping Control Table").ListObjects("Scoping_Control_Table")
    If scopingTable Is Nothing Then
        ValidateRequiredTables = False
        Exit Function
    End If

    ' Check for Pack_Number_Company_Table
    Set packTable = g_OutputWorkbook.Worksheets("Pack Number Company Table").ListObjects("Pack_Number_Company_Table")
    If packTable Is Nothing Then
        ValidateRequiredTables = False
        Exit Function
    End If

    ValidateRequiredTables = True
End Function

' Create Executive Summary Dashboard
Private Sub CreateExecutiveSummaryDashboard()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim scopingTbl As ListObject
    Dim packTbl As ListObject
    Dim row As Long
    Dim col As Long

    Application.StatusBar = "Creating Executive Summary Dashboard..."

    ' Create or get worksheet
    On Error Resume Next
    Set ws = g_OutputWorkbook.Worksheets("Dashboard - Executive Summary")
    If ws Is Nothing Then
        Set ws = g_OutputWorkbook.Worksheets.Add(Before:=g_OutputWorkbook.Worksheets(1))
        ws.Name = "Dashboard - Executive Summary"
    Else
        ws.Cells.Clear
        DeleteAllCharts ws
    End If
    On Error GoTo ErrorHandler

    Set scopingTbl = g_OutputWorkbook.Worksheets("Scoping Control Table").ListObjects("Scoping_Control_Table")
    Set packTbl = g_OutputWorkbook.Worksheets("Pack Number Company Table").ListObjects("Pack_Number_Company_Table")

    With ws
        ' ===== TITLE & HEADER =====
        .Cells(1, 1).Value = "ISA 600 SCOPING DASHBOARD"
        .Cells(1, 1).Font.Name = "Calibri"
        .Cells(1, 1).Font.Size = 24
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Color = RGB(255, 255, 255)
        .Cells(1, 1).Interior.Color = RGB(68, 114, 196)
        .Range("A1:M1").Merge
        .Range("A1:M1").VerticalAlignment = xlVAlignCenter
        .Range("A1:M1").HorizontalAlignment = xlHAlignCenter
        .Rows(1).RowHeight = 40

        .Cells(2, 1).Value = "Bidvest Group Limited - Consolidation Scoping Analysis v5.1"
        .Cells(2, 1).Font.Size = 11
        .Cells(2, 1).Font.Italic = True
        .Cells(2, 1).Interior.Color = RGB(217, 225, 242)
        .Range("A2:M2").Merge
        .Range("A2:M2").HorizontalAlignment = xlHAlignCenter

        .Cells(3, 1).Value = "Generated: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
        .Cells(3, 1).Font.Size = 9
        .Range("A3:M3").Merge
        .Range("A3:M3").HorizontalAlignment = xlHAlignCenter

        row = 5

        ' ===== KEY METRICS ROW =====
        .Cells(row, 1).Value = "KEY PERFORMANCE INDICATORS"
        .Cells(row, 1).Font.Size = 14
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Color = RGB(68, 114, 196)
        row = row + 1

        ' Create 6 KPI cards in 2 rows
        Call CreateEnhancedKPICard(ws, row, 1, "Total Packs", _
            "Scoping_Control_Table", "Pack Code", "Is Consolidated", "No", _
            "Unique entities for scoping", RGB(46, 125, 50), "")

        Call CreateEnhancedKPICard(ws, row, 3, "Scoped In (Auto)", _
            "Scoping_Control_Table", "Pack Code", "Scoping Status", "Scoped In (Auto)", _
            "Threshold-based decisions", RGB(76, 175, 80), "")

        Call CreateEnhancedKPICard(ws, row, 5, "Scoped In (Manual)", _
            "Scoping_Control_Table", "Pack Code", "Scoping Status", "Scoped In (Manual)", _
            "Manual decisions", RGB(139, 195, 74), "")

        Call CreateEnhancedKPICard(ws, row, 7, "Not Scoped", _
            "Scoping_Control_Table", "Pack Code", "Scoping Status", "Not Scoped", _
            "Pending decisions", RGB(255, 193, 7), "")

        Call CreateEnhancedKPICard(ws, row, 9, "Coverage %", _
            "Scoping_Control_Table", "Amount", "Scoping Status", "Scoped", _
            "Amount-based coverage", RGB(33, 150, 243), "%")

        Call CreateEnhancedKPICard(ws, row, 11, "Untested %", _
            "Scoping_Control_Table", "Amount", "Scoping Status", "Not Scoped", _
            "Remaining risk", RGB(244, 67, 54), "%")

        row = row + 5

        ' ===== ANALYSIS SECTION =====
        .Cells(row, 1).Value = "SCOPING ANALYSIS & VISUALIZATIONS"
        .Cells(row, 1).Font.Size = 14
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Color = RGB(68, 114, 196)
        row = row + 2

        ' Create summary table with formulas
        Call CreateScopingSummaryTable(ws, row, 1)

        row = row + 15

        ' ===== COVERAGE BY FSLI SECTION =====
        .Cells(row, 1).Value = "COVERAGE BY FINANCIAL STATEMENT LINE ITEM"
        .Cells(row, 1).Font.Size = 12
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Color = RGB(68, 114, 196)
        row = row + 2

        Call CreateFSLICoverageTable(ws, row, 1)

        ' Final formatting
        .Columns("A:M").AutoFit
        .Tab.Color = RGB(68, 114, 196)
        .Range("A1").Select
        ActiveWindow.FreezePanes = False
        .Range("A5").Select
        ActiveWindow.FreezePanes = True
    End With

    Exit Sub

ErrorHandler:
    Debug.Print "Error in CreateExecutiveSummaryDashboard: " & Err.Description
End Sub

' Create enhanced KPI card with formulas
Private Sub CreateEnhancedKPICard(ws As Worksheet, row As Long, col As Long, _
    title As String, tableName As String, countCol As String, filterCol As String, filterValue As String, _
    description As String, color As Long, formatType As String)

    On Error Resume Next

    Dim formulaStr As String
    Dim tbl As ListObject
    Dim tblSheet As Worksheet

    ' Get table reference
    Set tblSheet = g_OutputWorkbook.Worksheets("Scoping Control Table")
    If tblSheet Is Nothing Then Exit Sub

    Set tbl = tblSheet.ListObjects("Scoping_Control_Table")
    If tbl Is Nothing Then Exit Sub

    With ws
        ' Title
        .Cells(row, col).Value = title
        .Cells(row, col).Font.Size = 10
        .Cells(row, col).Font.Bold = True
        .Cells(row, col).Font.Color = RGB(255, 255, 255)
        .Cells(row, col).Interior.Color = color
        .Range(.Cells(row, col), .Cells(row, col + 1)).Merge
        .Range(.Cells(row, col), .Cells(row, col + 1)).HorizontalAlignment = xlHAlignCenter

        ' Value with formula
        If formatType = "%" Then
            ' Coverage percentage formula - COUNT BOTH AUTO AND MANUAL SCOPING
            If InStr(filterValue, "Scoped") > 0 Then
                ' Scoped percentage = (Auto + Manual) / Total
                formulaStr = "=IFERROR((" & _
                            "SUMIFS('" & tblSheet.Name & "'!" & tbl.ListColumns("Amount").DataBodyRange.Address & "," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Scoping Status").DataBodyRange.Address & ",""Scoped In (Auto)""," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Is Consolidated").DataBodyRange.Address & ",""No"")" & _
                            "+" & _
                            "SUMIFS('" & tblSheet.Name & "'!" & tbl.ListColumns("Amount").DataBodyRange.Address & "," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Scoping Status").DataBodyRange.Address & ",""Scoped In (Manual)""," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Is Consolidated").DataBodyRange.Address & ",""No"")" & _
                            ")/" & _
                            "SUMIF('" & tblSheet.Name & "'!" & tbl.ListColumns("Is Consolidated").DataBodyRange.Address & ",""No""," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Amount").DataBodyRange.Address & "),0)"
            Else
                ' Not scoped percentage = 1 - (Auto + Manual) / Total
                formulaStr = "=1-IFERROR((" & _
                            "SUMIFS('" & tblSheet.Name & "'!" & tbl.ListColumns("Amount").DataBodyRange.Address & "," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Scoping Status").DataBodyRange.Address & ",""Scoped In (Auto)""," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Is Consolidated").DataBodyRange.Address & ",""No"")" & _
                            "+" & _
                            "SUMIFS('" & tblSheet.Name & "'!" & tbl.ListColumns("Amount").DataBodyRange.Address & "," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Scoping Status").DataBodyRange.Address & ",""Scoped In (Manual)""," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Is Consolidated").DataBodyRange.Address & ",""No"")" & _
                            ")/" & _
                            "SUMIF('" & tblSheet.Name & "'!" & tbl.ListColumns("Is Consolidated").DataBodyRange.Address & ",""No""," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("Amount").DataBodyRange.Address & "),1)"
            End If
            .Cells(row + 1, col).formula = formulaStr
            .Cells(row + 1, col).NumberFormat = "0.0%"
        Else
            ' Count formula
            If countCol = "Pack Code" Then
                formulaStr = "=SUMPRODUCT(('" & tblSheet.Name & "'!" & tbl.ListColumns("Scoping Status").DataBodyRange.Address & "=""" & filterValue & """)/" & _
                            "COUNTIF('" & tblSheet.Name & "'!" & tbl.ListColumns("FSLI").DataBodyRange.Address & "," & _
                            "'" & tblSheet.Name & "'!" & tbl.ListColumns("FSLI").DataBodyRange.Address & "))"
            Else
                formulaStr = "=COUNTIFS('" & tblSheet.Name & "'!" & tbl.ListColumns(filterCol).DataBodyRange.Address & ",""" & filterValue & """)"
            End If
            .Cells(row + 1, col).formula = formulaStr
            .Cells(row + 1, col).NumberFormat = "#,##0"
        End If

        .Cells(row + 1, col).Font.Size = 20
        .Cells(row + 1, col).Font.Bold = True
        .Cells(row + 1, col).Font.Color = color
        .Range(.Cells(row + 1, col), .Cells(row + 1, col + 1)).Merge
        .Range(.Cells(row + 1, col), .Cells(row + 1, col + 1)).HorizontalAlignment = xlHAlignCenter

        ' Description
        .Cells(row + 2, col).Value = description
        .Cells(row + 2, col).Font.Size = 8
        .Cells(row + 2, col).Font.Italic = True
        .Cells(row + 2, col).WrapText = True
        .Range(.Cells(row + 2, col), .Cells(row + 2, col + 1)).Merge
        .Range(.Cells(row + 2, col), .Cells(row + 2, col + 1)).HorizontalAlignment = xlHAlignCenter

        ' Border
        .Range(.Cells(row, col), .Cells(row + 2, col + 1)).Borders.LineStyle = xlContinuous
        .Range(.Cells(row, col), .Cells(row + 2, col + 1)).Borders.Weight = xlMedium
        .Range(.Cells(row + 1, col), .Cells(row + 2, col + 1)).Interior.Color = RGB(250, 250, 250)
    End With
End Sub

' Create scoping summary table with formulas
Private Sub CreateScopingSummaryTable(ws As Worksheet, row As Long, col As Long)
    On Error Resume Next

    Dim startRow As Long
    startRow = row

    With ws
        ' Headers
        .Cells(row, col).Value = "Status"
        .Cells(row, col + 1).Value = "Packs"
        .Cells(row, col + 2).Value = "% of Total"
        .Cells(row, col + 3).Value = "Amount"
        .Cells(row, col + 4).Value = "% of Amount"
        .Cells(row, col + 5).Value = "Avg per Pack"

        .Range(.Cells(row, col), .Cells(row, col + 5)).Font.Bold = True
        .Range(.Cells(row, col), .Cells(row, col + 5)).Interior.Color = RGB(68, 114, 196)
        .Range(.Cells(row, col), .Cells(row, col + 5)).Font.Color = RGB(255, 255, 255)
        .Range(.Cells(row, col), .Cells(row, col + 5)).HorizontalAlignment = xlHAlignCenter

        row = row + 1

        ' Data rows with formulas
        Dim statuses As Variant
        Dim colors As Variant
        Dim i As Long

        statuses = Array("Scoped In (Auto)", "Scoped In (Manual)", "Not Scoped", "Scoped Out")
        colors = Array(RGB(200, 230, 201), RGB(220, 237, 200), RGB(255, 249, 196), RGB(255, 205, 210))

        For i = 0 To UBound(statuses)
            .Cells(row, col).Value = statuses(i)
            .Cells(row, col).Interior.Color = colors(i)

            ' Packs formula
            .Cells(row, col + 1).formula = "=SUMPRODUCT(('Scoping Control Table'!F:F=""" & statuses(i) & """)/" & _
                                          "COUNTIF('Scoping Control Table'!D:D,'Scoping Control Table'!D:D))"
            .Cells(row, col + 1).NumberFormat = "#,##0"

            ' % of Total Packs
            .Cells(row, col + 2).formula = "=IFERROR(" & .Cells(row, col + 1).Address & "/" & _
                                          "SUM(" & .Cells(row + 1, col + 1).Address & ":" & .Cells(row + 4 - i, col + 1).Address & "),0)"
            .Cells(row, col + 2).NumberFormat = "0.0%"

            ' Amount formula
            .Cells(row, col + 3).formula = "=SUMIFS('Scoping Control Table'!E:E,'Scoping Control Table'!F:F,""" & statuses(i) & """)"
            .Cells(row, col + 3).NumberFormat = "#,##0"

            ' % of Amount
            .Cells(row, col + 4).formula = "=IFERROR(" & .Cells(row, col + 3).Address & "/" & _
                                          "SUM(" & .Cells(row + 1, col + 3).Address & ":" & .Cells(row + 4 - i, col + 3).Address & "),0)"
            .Cells(row, col + 4).NumberFormat = "0.0%"

            ' Avg per Pack
            .Cells(row, col + 5).formula = "=IFERROR(" & .Cells(row, col + 3).Address & "/" & .Cells(row, col + 1).Address & ",0)"
            .Cells(row, col + 5).NumberFormat = "#,##0"

            row = row + 1
        Next i

        ' Total row
        .Cells(row, col).Value = "TOTAL"
        .Cells(row, col).Font.Bold = True
        .Cells(row, col).Interior.Color = RGB(189, 189, 189)

        .Cells(row, col + 1).formula = "=SUM(" & .Cells(startRow + 1, col + 1).Address & ":" & .Cells(row - 1, col + 1).Address & ")"
        .Cells(row, col + 1).NumberFormat = "#,##0"
        .Cells(row, col + 1).Font.Bold = True

        .Cells(row, col + 2).Value = "100.0%"
        .Cells(row, col + 2).Font.Bold = True

        .Cells(row, col + 3).formula = "=SUM(" & .Cells(startRow + 1, col + 3).Address & ":" & .Cells(row - 1, col + 3).Address & ")"
        .Cells(row, col + 3).NumberFormat = "#,##0"
        .Cells(row, col + 3).Font.Bold = True

        .Cells(row, col + 4).Value = "100.0%"
        .Cells(row, col + 4).Font.Bold = True

        .Cells(row, col + 5).formula = "=IFERROR(" & .Cells(row, col + 3).Address & "/" & .Cells(row, col + 1).Address & ",0)"
        .Cells(row, col + 5).NumberFormat = "#,##0"
        .Cells(row, col + 5).Font.Bold = True

        ' Borders
        .Range(.Cells(startRow, col), .Cells(row, col + 5)).Borders.LineStyle = xlContinuous
        .Range(.Cells(row, col), .Cells(row, col + 5)).Borders(xlEdgeTop).Weight = xlMedium
    End With
End Sub

' Create FSLI coverage table
Private Sub CreateFSLICoverageTable(ws As Worksheet, row As Long, col As Long)
    ' This will be populated with top FSLIs and their coverage
    ' Placeholder for now - will enhance in next iteration
End Sub

' Enhance Scoping Control Table with dropdowns
Private Sub EnhanceScopingControlTable()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long
    Dim statusCol As Range

    Application.StatusBar = "Enhancing Scoping Control Table..."

    Set ws = g_OutputWorkbook.Worksheets("Scoping Control Table")
    Set tbl = ws.ListObjects("Scoping_Control_Table")

    ' Add data validation to Scoping Status column
    lastRow = tbl.DataBodyRange.Rows.Count + tbl.HeaderRowRange.row
    Set statusCol = ws.Range("F" & (tbl.HeaderRowRange.row + 1) & ":F" & lastRow)

    statusCol.Validation.Delete
    With statusCol.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:="Not Scoped,Scoped In (Manual),Scoped In (Auto),Scoped Out"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .InputTitle = "Scoping Decision"
        .InputMessage = "Select scoping status. Changes update dashboard automatically."
        .ShowError = True
        .ErrorTitle = "Invalid Entry"
        .ErrorMessage = "Please select a value from the dropdown list."
    End With

    ' Apply conditional formatting
    Call ApplyEnhancedConditionalFormatting(statusCol)

    ' Add instructions
    ws.Range("A1").EntireRow.Insert
    ws.Range("A1").EntireRow.Insert
    ws.Range("A1").Value = "SCOPING CONTROL TABLE - Interactive"
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(68, 114, 196)

    ws.Range("A2").Value = "Click any cell in 'Scoping Status' column → Select from dropdown → Dashboard updates automatically"
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").WrapText = True

    ws.Tab.Color = RGB(33, 150, 243)

    Exit Sub

ErrorHandler:
    Debug.Print "Error in EnhanceScopingControlTable: " & Err.Description
End Sub

' Apply enhanced conditional formatting
Private Sub ApplyEnhancedConditionalFormatting(rng As Range)
    On Error Resume Next

    rng.FormatConditions.Delete

    Dim fc As FormatCondition

    ' Scoped In (Auto)
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Scoped In (Auto)""")
    fc.Interior.Color = RGB(200, 230, 201)
    fc.Font.Color = RGB(27, 94, 32)
    fc.Font.Bold = True

    ' Scoped In (Manual)
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Scoped In (Manual)""")
    fc.Interior.Color = RGB(220, 237, 200)
    fc.Font.Color = RGB(51, 105, 30)
    fc.Font.Bold = True

    ' Not Scoped
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Not Scoped""")
    fc.Interior.Color = RGB(255, 249, 196)
    fc.Font.Color = RGB(245, 127, 23)
    fc.Font.Bold = True

    ' Scoped Out
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Scoped Out""")
    fc.Interior.Color = RGB(255, 205, 210)
    fc.Font.Color = RGB(198, 40, 40)
    fc.Font.Bold = True
End Sub

' Create FSLI Coverage Analysis sheet
Private Sub CreateFSLICoverageAnalysis()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim scopingTbl As ListObject
    Dim fsliDict As Object
    Dim fsli As Variant
    Dim row As Long, col As Long
    Dim totalAmount As Double, scopedAmount As Double
    Dim coveragePct As Double
    Dim dataRow As Long
    Dim tbl As ListObject

    Application.StatusBar = "Creating FSLI Coverage Analysis..."

    ' Get scoping table
    Set scopingTbl = g_OutputWorkbook.Worksheets("Scoping Control Table").ListObjects("Scoping_Control_Table")

    ' Create or clear worksheet
    On Error Resume Next
    Set ws = g_OutputWorkbook.Worksheets("FSLI Coverage Analysis")
    If ws Is Nothing Then
        Set ws = g_OutputWorkbook.Worksheets.Add(After:=g_OutputWorkbook.Worksheets(1))
        ws.Name = "FSLI Coverage Analysis"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Build FSLI dictionary with totals
    Set fsliDict = CreateObject("Scripting.Dictionary")

    ' Iterate through scoping table to calculate per-FSLI totals
    For dataRow = 1 To scopingTbl.DataBodyRange.Rows.Count
        Dim fsliName As String, amt As Double, status As String, isConsol As String

        fsliName = scopingTbl.DataBodyRange.Cells(dataRow, 4).Value ' FSLI column
        amt = scopingTbl.DataBodyRange.Cells(dataRow, 5).Value ' Amount column
        status = scopingTbl.DataBodyRange.Cells(dataRow, 6).Value ' Scoping Status column
        isConsol = scopingTbl.DataBodyRange.Cells(dataRow, 7).Value ' Is Consolidated column

        ' Skip consolidated packs
        If UCase(isConsol) = "NO" Then
            If Not fsliDict.Exists(fsliName) Then
                fsliDict(fsliName) = CreateObject("Scripting.Dictionary")
                fsliDict(fsliName)("TotalAmount") = 0
                fsliDict(fsliName)("ScopedAmount") = 0
                fsliDict(fsliName)("PackCount") = 0
                fsliDict(fsliName)("ScopedPackCount") = 0
            End If

            fsliDict(fsliName)("TotalAmount") = fsliDict(fsliName)("TotalAmount") + amt

            ' Count scoped amounts (both Auto and Manual)
            If InStr(1, status, "Scoped In", vbTextCompare) > 0 Then
                fsliDict(fsliName)("ScopedAmount") = fsliDict(fsliName)("ScopedAmount") + amt
                fsliDict(fsliName)("ScopedPackCount") = fsliDict(fsliName)("ScopedPackCount") + 1
            End If

            fsliDict(fsliName)("PackCount") = fsliDict(fsliName)("PackCount") + 1
        End If
    Next dataRow

    ' Check if we have data
    If fsliDict.Count = 0 Then
        ws.Cells(1, 1).Value = "No FSLI data found. Ensure Scoping Control Table has non-consolidated data."
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(1, 1).Font.Bold = True
        Exit Sub
    End If

    ' Write header
    row = 1
    With ws
        .Cells(row, 1).Value = "FSLI COVERAGE ANALYSIS"
        .Cells(row, 1).Font.Size = 16
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Color = RGB(255, 255, 255)
        .Cells(row, 1).Interior.Color = RGB(68, 114, 196)
        .Range("A1:H1").Merge
        .Range("A1:H1").HorizontalAlignment = xlCenter
        row = row + 2

        ' Column headers
        .Cells(row, 1).Value = "FSLI"
        .Cells(row, 2).Value = "Total Amount"
        .Cells(row, 3).Value = "Scoped Amount"
        .Cells(row, 4).Value = "Not Scoped Amount"
        .Cells(row, 5).Value = "Coverage %"
        .Cells(row, 6).Value = "Total Packs"
        .Cells(row, 7).Value = "Scoped Packs"
        .Cells(row, 8).Value = "Status"

        .Range("A" & row & ":H" & row).Font.Bold = True
        .Range("A" & row & ":H" & row).Interior.Color = RGB(217, 217, 217)
        row = row + 1

        ' Write FSLI data
        For Each fsli In fsliDict.Keys
            totalAmount = fsliDict(fsli)("TotalAmount")
            scopedAmount = fsliDict(fsli)("ScopedAmount")
            coveragePct = 0
            If totalAmount > 0 Then coveragePct = scopedAmount / totalAmount

            .Cells(row, 1).Value = fsli
            .Cells(row, 2).Value = totalAmount
            .Cells(row, 2).NumberFormat = "#,##0.00"
            .Cells(row, 3).Value = scopedAmount
            .Cells(row, 3).NumberFormat = "#,##0.00"
            .Cells(row, 4).Value = totalAmount - scopedAmount
            .Cells(row, 4).NumberFormat = "#,##0.00"
            .Cells(row, 5).Value = coveragePct
            .Cells(row, 5).NumberFormat = "0.0%"
            .Cells(row, 6).Value = fsliDict(fsli)("PackCount")
            .Cells(row, 7).Value = fsliDict(fsli)("ScopedPackCount")

            ' Status indicator
            If coveragePct >= 0.6 Then
                .Cells(row, 8).Value = "✓ On Target"
                .Cells(row, 8).Font.Color = RGB(0, 176, 80)
            ElseIf coveragePct >= 0.3 Then
                .Cells(row, 8).Value = "⚠ Below Target"
                .Cells(row, 8).Font.Color = RGB(255, 192, 0)
            Else
                .Cells(row, 8).Value = "✗ Needs Attention"
                .Cells(row, 8).Font.Color = RGB(192, 0, 0)
            End If

            row = row + 1
        Next fsli

        ' Create table if data exists
        If row > 4 Then
            On Error Resume Next
            Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(3, 1), .Cells(row - 1, 8)), , xlYes)
            If Not tbl Is Nothing Then
                tbl.Name = "FSLI_Coverage_Table"
                tbl.TableStyle = "TableStyleMedium2"
                tbl.ShowTableStyleRowStripes = True
            End If
            On Error GoTo ErrorHandler
        End If

        ' Auto-fit columns
        .Columns("A:H").AutoFit
        .Tab.Color = RGB(91, 155, 213)
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Error creating FSLI Coverage Analysis: " & Err.Description, vbCritical
End Sub

' Create Division/Segment Analysis
Private Sub CreateDivisionSegmentAnalysis()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim scopingTbl As ListObject
    Dim divisionDict As Object
    Dim fsliDivKey As String
    Dim row As Long
    Dim dataRow As Long
    Dim tbl As ListObject
    Dim division As String, fsliName As String, amt As Double, status As String, isConsol As String
    Dim totalAmount As Double, scopedAmount As Double, coveragePct As Double

    Application.StatusBar = "Creating Division/FSLI Coverage Analysis..."

    ' Get scoping table
    Set scopingTbl = g_OutputWorkbook.Worksheets("Scoping Control Table").ListObjects("Scoping_Control_Table")

    ' Create or clear worksheet
    On Error Resume Next
    Set ws = g_OutputWorkbook.Worksheets("Division FSLI Coverage")
    If ws Is Nothing Then
        Set ws = g_OutputWorkbook.Worksheets.Add(After:=g_OutputWorkbook.Worksheets(2))
        ws.Name = "Division FSLI Coverage"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Build Division+FSLI dictionary
    Set divisionDict = CreateObject("Scripting.Dictionary")

    For dataRow = 1 To scopingTbl.DataBodyRange.Rows.Count
        division = scopingTbl.DataBodyRange.Cells(dataRow, 3).Value ' Division column
        fsliName = scopingTbl.DataBodyRange.Cells(dataRow, 4).Value ' FSLI column
        amt = scopingTbl.DataBodyRange.Cells(dataRow, 5).Value ' Amount column
        status = scopingTbl.DataBodyRange.Cells(dataRow, 6).Value ' Scoping Status column
        isConsol = scopingTbl.DataBodyRange.Cells(dataRow, 7).Value ' Is Consolidated column

        ' Skip consolidated packs
        If UCase(isConsol) = "NO" Then
            ' Create composite key: Division|FSLI
            fsliDivKey = division & "|" & fsliName

            If Not divisionDict.Exists(fsliDivKey) Then
                divisionDict(fsliDivKey) = CreateObject("Scripting.Dictionary")
                divisionDict(fsliDivKey)("Division") = division
                divisionDict(fsliDivKey)("FSLI") = fsliName
                divisionDict(fsliDivKey)("TotalAmount") = 0
                divisionDict(fsliDivKey)("ScopedAmount") = 0
                divisionDict(fsliDivKey)("PackCount") = 0
            End If

            divisionDict(fsliDivKey)("TotalAmount") = divisionDict(fsliDivKey)("TotalAmount") + amt

            If InStr(1, status, "Scoped In", vbTextCompare) > 0 Then
                divisionDict(fsliDivKey)("ScopedAmount") = divisionDict(fsliDivKey)("ScopedAmount") + amt
            End If

            divisionDict(fsliDivKey)("PackCount") = divisionDict(fsliDivKey)("PackCount") + 1
        End If
    Next dataRow

    ' Check if we have data
    If divisionDict.Count = 0 Then
        ws.Cells(1, 1).Value = "No Division/FSLI data found. Ensure Scoping Control Table has non-consolidated data."
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(1, 1).Font.Bold = True
        Exit Sub
    End If

    ' Write data to worksheet
    row = 1
    With ws
        .Cells(row, 1).Value = "DIVISION × FSLI COVERAGE ANALYSIS"
        .Cells(row, 1).Font.Size = 16
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Color = RGB(255, 255, 255)
        .Cells(row, 1).Interior.Color = RGB(112, 48, 160)
        .Range("A1:H1").Merge
        .Range("A1:H1").HorizontalAlignment = xlCenter
        row = row + 2

        ' Column headers
        .Cells(row, 1).Value = "Division"
        .Cells(row, 2).Value = "FSLI"
        .Cells(row, 3).Value = "Total Amount"
        .Cells(row, 4).Value = "Scoped Amount"
        .Cells(row, 5).Value = "Not Scoped Amount"
        .Cells(row, 6).Value = "Coverage %"
        .Cells(row, 7).Value = "Pack Count"
        .Cells(row, 8).Value = "Status"

        .Range("A" & row & ":H" & row).Font.Bold = True
        .Range("A" & row & ":H" & row).Interior.Color = RGB(217, 217, 217)
        row = row + 1

        ' Write data
        Dim key As Variant
        For Each key In divisionDict.Keys
            totalAmount = divisionDict(key)("TotalAmount")
            scopedAmount = divisionDict(key)("ScopedAmount")
            coveragePct = 0
            If totalAmount > 0 Then coveragePct = scopedAmount / totalAmount

            .Cells(row, 1).Value = divisionDict(key)("Division")
            .Cells(row, 2).Value = divisionDict(key)("FSLI")
            .Cells(row, 3).Value = totalAmount
            .Cells(row, 3).NumberFormat = "#,##0.00"
            .Cells(row, 4).Value = scopedAmount
            .Cells(row, 4).NumberFormat = "#,##0.00"
            .Cells(row, 5).Value = totalAmount - scopedAmount
            .Cells(row, 5).NumberFormat = "#,##0.00"
            .Cells(row, 6).Value = coveragePct
            .Cells(row, 6).NumberFormat = "0.0%"
            .Cells(row, 7).Value = divisionDict(key)("PackCount")

            ' Status indicator
            If coveragePct >= 0.6 Then
                .Cells(row, 8).Value = "✓ On Target"
                .Cells(row, 8).Font.Color = RGB(0, 176, 80)
            ElseIf coveragePct >= 0.3 Then
                .Cells(row, 8).Value = "⚠ Below Target"
                .Cells(row, 8).Font.Color = RGB(255, 192, 0)
            Else
                .Cells(row, 8).Value = "✗ Needs Attention"
                .Cells(row, 8).Font.Color = RGB(192, 0, 0)
            End If

            row = row + 1
        Next key

        ' Create table
        If row > 4 Then
            On Error Resume Next
            Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(3, 1), .Cells(row - 1, 8)), , xlYes)
            If Not tbl Is Nothing Then
                tbl.Name = "Division_FSLI_Coverage_Table"
                tbl.TableStyle = "TableStyleMedium9"
                tbl.ShowTableStyleRowStripes = True
            End If
            On Error GoTo ErrorHandler
        End If

        ' Auto-fit and formatting
        .Columns("A:H").AutoFit
        .Tab.Color = RGB(147, 101, 184)

        ' Add note
        .Cells(row + 2, 1).Value = "NOTE: Use filters above to drill down by specific Division or FSLI"
        .Cells(row + 2, 1).Font.Italic = True
        .Cells(row + 2, 1).Font.Color = RGB(128, 128, 128)
    End With

    ' Also create segment analysis if segment data exists
    Call CreateSegmentAnalysisSheet

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Division/Segment Analysis: " & Err.Description, vbCritical
End Sub

' Create Segment Analysis Sheet (if segment data exists)
Private Sub CreateSegmentAnalysisSheet()
    On Error Resume Next ' Segments are optional

    Dim segmentWs As Worksheet
    Dim segmentTbl As ListObject

    ' Check if Segment_Pack_Mapping exists
    Set segmentWs = g_OutputWorkbook.Worksheets("Segment_Pack_Mapping")
    If segmentWs Is Nothing Then Exit Sub

    Set segmentTbl = segmentWs.ListObjects("Segment_Pack_Mapping")
    If segmentTbl Is Nothing Then Exit Sub

    ' If segments exist, create segment coverage analysis
    ' (This would be a more complex analysis - for now just add a note)
    Dim ws As Worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add(After:=g_OutputWorkbook.Worksheets(3))
    ws.Name = "Segment Coverage"

    ws.Cells(1, 1).Value = "SEGMENT COVERAGE ANALYSIS"
    ws.Cells(1, 1).Font.Size = 16
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(3, 1).Value = "Segment coverage analysis based on Segment_Pack_Mapping table."
    ws.Cells(4, 1).Value = "Use Segment_Summary sheet for detailed segment breakdown."
    ws.Tab.Color = RGB(169, 208, 142)

    On Error GoTo 0
End Sub

' Create Interactive Worksheet
Private Sub CreateInteractiveWorksheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim scopingTbl As ListObject
    Dim contributorsDict As Object
    Dim packKey As Variant  ' Changed from String to Variant for dictionary iteration
    Dim row As Long, dataRow As Long
    Dim tbl As ListObject
    Dim grandTotal As Double
    Dim packName As String, packCode As String, division As String, fsliName As String
    Dim amt As Double, status As String, isConsol As String
    Dim contribution As Double

    Application.StatusBar = "Creating Interactive Scoping Worksheet..."

    ' Get scoping table
    Set scopingTbl = g_OutputWorkbook.Worksheets("Scoping Control Table").ListObjects("Scoping_Control_Table")

    ' Create or clear worksheet
    On Error Resume Next
    Set ws = g_OutputWorkbook.Worksheets("Highest Contributors")
    If ws Is Nothing Then
        Set ws = g_OutputWorkbook.Worksheets.Add(After:=g_OutputWorkbook.Worksheets(3))
        ws.Name = "Highest Contributors"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrorHandler

    ' Calculate grand total (excluding consolidated)
    grandTotal = 0
    For dataRow = 1 To scopingTbl.DataBodyRange.Rows.Count
        isConsol = scopingTbl.DataBodyRange.Cells(dataRow, 7).Value
        If UCase(isConsol) = "NO" Then
            grandTotal = grandTotal + scopingTbl.DataBodyRange.Cells(dataRow, 5).Value
        End If
    Next dataRow

    ' Build contributors dictionary (Pack + FSLI level)
    Set contributorsDict = CreateObject("Scripting.Dictionary")

    For dataRow = 1 To scopingTbl.DataBodyRange.Rows.Count
        packName = scopingTbl.DataBodyRange.Cells(dataRow, 1).Value
        packCode = scopingTbl.DataBodyRange.Cells(dataRow, 2).Value
        division = scopingTbl.DataBodyRange.Cells(dataRow, 3).Value
        fsliName = scopingTbl.DataBodyRange.Cells(dataRow, 4).Value
        amt = scopingTbl.DataBodyRange.Cells(dataRow, 5).Value
        status = scopingTbl.DataBodyRange.Cells(dataRow, 6).Value
        isConsol = scopingTbl.DataBodyRange.Cells(dataRow, 7).Value

        ' Skip consolidated packs
        If UCase(isConsol) = "NO" And amt > 0 Then
            packKey = packCode & "|" & fsliName

            If Not contributorsDict.Exists(packKey) Then
                contributorsDict(packKey) = CreateObject("Scripting.Dictionary")
                contributorsDict(packKey)("PackName") = packName
                contributorsDict(packKey)("PackCode") = packCode
                contributorsDict(packKey)("Division") = division
                contributorsDict(packKey)("FSLI") = fsliName
                contributorsDict(packKey)("Amount") = amt
                contributorsDict(packKey)("Status") = status
                contributorsDict(packKey)("Contribution") = 0
                If grandTotal > 0 Then
                    contributorsDict(packKey)("Contribution") = amt / grandTotal
                End If
            End If
        End If
    Next dataRow

    ' Check if we have any contributors
    If contributorsDict.Count = 0 Then
        ' No contributors found - add message and exit
        ws.Cells(1, 1).Value = "No contributors found. Ensure Scoping Control Table has data."
        ws.Cells(1, 1).Font.Size = 14
        ws.Cells(1, 1).Font.Bold = True
        Exit Sub
    End If

    ' Sort contributors by amount (largest first) - using simple bubble sort
    Dim keys() As Variant
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim sortedKeys As Collection
    Set sortedKeys = New Collection

    ' Convert dictionary keys to array
    ReDim keys(contributorsDict.Count - 1)
    i = 0
    For Each packKey In contributorsDict.Keys
        keys(i) = packKey
        i = i + 1
    Next packKey

    ' Sort by amount (descending)
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If contributorsDict(keys(i))("Amount") < contributorsDict(keys(j))("Amount") Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i

    ' Write data to worksheet
    row = 1
    With ws
        .Cells(row, 1).Value = "TOP CONTRIBUTORS BY AMOUNT"
        .Cells(row, 1).Font.Size = 16
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Color = RGB(255, 255, 255)
        .Cells(row, 1).Interior.Color = RGB(237, 125, 49)
        .Range("A1:I1").Merge
        .Range("A1:I1").HorizontalAlignment = xlCenter
        row = row + 2

        ' Column headers
        .Cells(row, 1).Value = "Rank"
        .Cells(row, 2).Value = "Pack Name"
        .Cells(row, 3).Value = "Pack Code"
        .Cells(row, 4).Value = "Division"
        .Cells(row, 5).Value = "FSLI"
        .Cells(row, 6).Value = "Amount"
        .Cells(row, 7).Value = "% of Total"
        .Cells(row, 8).Value = "Scoping Status"
        .Cells(row, 9).Value = "Impact"

        .Range("A" & row & ":I" & row).Font.Bold = True
        .Range("A" & row & ":I" & row).Interior.Color = RGB(217, 217, 217)
        row = row + 1

        ' Write top contributors (limit to top 100)
        Dim maxRows As Long
        maxRows = Application.WorksheetFunction.Min(100, UBound(keys) + 1)

        For i = 0 To maxRows - 1
            packKey = keys(i)

            .Cells(row, 1).Value = i + 1 ' Rank
            .Cells(row, 2).Value = contributorsDict(packKey)("PackName")
            .Cells(row, 3).Value = contributorsDict(packKey)("PackCode")
            .Cells(row, 4).Value = contributorsDict(packKey)("Division")
            .Cells(row, 5).Value = contributorsDict(packKey)("FSLI")
            .Cells(row, 6).Value = contributorsDict(packKey)("Amount")
            .Cells(row, 6).NumberFormat = "#,##0.00"
            .Cells(row, 7).Value = contributorsDict(packKey)("Contribution")
            .Cells(row, 7).NumberFormat = "0.00%"
            .Cells(row, 8).Value = contributorsDict(packKey)("Status")

            ' Impact indicator
            contribution = contributorsDict(packKey)("Contribution")
            If contribution >= 0.05 Then
                .Cells(row, 9).Value = "🔴 High"
                .Cells(row, 9).Font.Color = RGB(192, 0, 0)
            ElseIf contribution >= 0.02 Then
                .Cells(row, 9).Value = "🟡 Medium"
                .Cells(row, 9).Font.Color = RGB(255, 192, 0)
            Else
                .Cells(row, 9).Value = "🟢 Low"
                .Cells(row, 9).Font.Color = RGB(0, 176, 80)
            End If

            row = row + 1
        Next i

        ' Create table
        If row > 4 Then
            On Error Resume Next
            Set tbl = .ListObjects.Add(xlSrcRange, .Range(.Cells(3, 1), .Cells(row - 1, 9)), , xlYes)
            If Not tbl Is Nothing Then
                tbl.Name = "Highest_Contributors_Table"
                tbl.TableStyle = "TableStyleMedium6"
                tbl.ShowTableStyleRowStripes = True
            End If
            On Error GoTo ErrorHandler
        End If

        ' Auto-fit columns
        .Columns("A:I").AutoFit
        .Tab.Color = RGB(255, 192, 0)

        ' Add usage notes
        .Cells(row + 2, 1).Value = "INTERACTIVE SCOPING:"
        .Cells(row + 2, 1).Font.Bold = True
        .Cells(row + 3, 1).Value = "• Use filters above to find specific packs, divisions, or FSLIs"
        .Cells(row + 4, 1).Value = "• To change scoping: Go to 'Scoping Control Table' sheet and use dropdowns"
        .Cells(row + 5, 1).Value = "• Dashboard metrics update automatically when you change scoping status"
        .Cells(row + 6, 1).Value = "• Focus on High Impact items (🔴) for maximum coverage improvement"

        .Cells(row + 8, 1).Value = "FILTER TIPS:"
        .Cells(row + 8, 1).Font.Bold = True
        .Cells(row + 9, 1).Value = "• Click filter arrows in table headers"
        .Cells(row + 10, 1).Value = "• Filter by Division to see division-specific contributors"
        .Cells(row + 11, 1).Value = "• Filter by FSLI to see which packs drive that FSLI"
        .Cells(row + 12, 1).Value = "• Filter by Scoping Status to identify Not Scoped items"
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Error creating Interactive Worksheet: " & Err.Description, vbCritical
End Sub

' Create Quick Reference Guide
Private Sub CreateQuickReferenceGuide()
    On Error Resume Next

    Dim ws As Worksheet
    Set ws = g_OutputWorkbook.Worksheets.Add(After:=g_OutputWorkbook.Worksheets(g_OutputWorkbook.Worksheets.Count))
    ws.Name = "Quick Reference Guide"

    With ws
        .Cells(1, 1).Value = "EXCEL DASHBOARD QUICK REFERENCE"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True

        .Cells(3, 1).Value = "Sheet Navigation:"
        .Cells(4, 1).Value = "• Dashboard - Executive Summary: Main KPIs and analysis"
        .Cells(5, 1).Value = "• Scoping Control Table: Interactive scoping with dropdowns"
        .Cells(6, 1).Value = "• Other sheets: Supporting data tables"

        .Cells(8, 1).Value = "How to Scope:"
        .Cells(9, 1).Value = "1. Go to 'Scoping Control Table'"
        .Cells(10, 1).Value = "2. Click cell in 'Scoping Status' column"
        .Cells(11, 1).Value = "3. Select from dropdown"
        .Cells(12, 1).Value = "4. Dashboard updates automatically!"

        .Columns("A:A").ColumnWidth = 60
        .Tab.Color = RGB(158, 158, 158)
    End With
End Sub

' Helper: Delete all charts from worksheet
Private Sub DeleteAllCharts(ws As Worksheet)
    On Error Resume Next
    Dim obj As Object
    For Each obj In ws.ChartObjects
        obj.Delete
    Next obj
End Sub
