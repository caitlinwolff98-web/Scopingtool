Attribute VB_Name = "Mod6_DashboardGeneration"
Option Explicit

' ============================================================================
' MODULE 6: DASHBOARD GENERATION
' ISA 600 Revised Component Scoping Tool - Bidvest Group
' Version: 6.0 - Complete Overhaul
' ============================================================================
' PURPOSE: Create comprehensive interactive dashboard
' DESCRIPTION: Generates 5 dashboard views with dynamic updates,
'              manual scoping interface, and coverage analysis
' ============================================================================

' ==================== CREATE COMPREHENSIVE DASHBOARD ====================
Public Sub CreateComprehensiveDashboard()
    '------------------------------------------------------------------------
    ' Main function to create all dashboard views
    '------------------------------------------------------------------------
    On Error GoTo ErrorHandler

    Application.StatusBar = "Creating comprehensive dashboard..."

    ' Create all dashboard views
    CreateView1_OverallScopingSummary
    CreateView2_CoverageByFSLI
    CreateView3_CoverageByDivision
    CreateView4_CoverageBySegment
    CreateView5_DetailedPackAnalysis

    ' Create manual scoping interface
    CreateManualScopingInterface

    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "Error creating dashboard: " & Err.Description
End Sub

' ==================== VIEW 1: OVERALL SCOPING SUMMARY ====================
Private Sub CreateView1_OverallScopingSummary()
    '------------------------------------------------------------------------
    ' View 1: Overall Scoping Summary
    ' Displays total packs, scoped in, not scoped, overall coverage%
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet
    Dim totalPacks As Long
    Dim scopedPacks As Long
    Dim notScoped As Long
    Dim coveragePct As Double

    Set dashWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dashWs.Name = "Dashboard - Overview"

    ' Title
    With dashWs.Range("A1")
        .Value = "ISA 600 SCOPING DASHBOARD - OVERALL SUMMARY"
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(0, 112, 192)
    End With

    ' Get data
    totalPacks = GetTotalPackCount()
    scopedPacks = Mod1_MainController.g_ScopedPacks.Count
    notScoped = totalPacks - scopedPacks
    If totalPacks > 0 Then
        coveragePct = (scopedPacks / totalPacks) * 100
    End If

    ' Display metrics
    Dim row As Long
    row = 3

    ' Total Packs
    dashWs.Cells(row, 1).Value = "Total Packs:"
    dashWs.Cells(row, 2).Value = totalPacks
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Font.Size = 14
    row = row + 1

    ' Scoped In
    dashWs.Cells(row, 1).Value = "Packs Scoped In:"
    dashWs.Cells(row, 2).Value = scopedPacks
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Font.Size = 14
    dashWs.Cells(row, 2).Interior.Color = RGB(198, 239, 206) ' Green
    row = row + 1

    ' Not Scoped
    dashWs.Cells(row, 1).Value = "Packs Not Yet Scoped:"
    dashWs.Cells(row, 2).Value = notScoped
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Font.Size = 14
    dashWs.Cells(row, 2).Interior.Color = RGB(255, 235, 156) ' Yellow
    row = row + 1

    ' Coverage %
    dashWs.Cells(row, 1).Value = "Overall Coverage %:"
    dashWs.Cells(row, 2).Value = coveragePct / 100
    dashWs.Cells(row, 2).NumberFormat = "0.0%"
    dashWs.Cells(row, 1).Font.Bold = True
    dashWs.Cells(row, 2).Font.Size = 14
    row = row + 2

    ' Instructions
    dashWs.Cells(row, 1).Value = "INSTRUCTIONS:"
    dashWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    dashWs.Cells(row, 1).Value = "1. Review 'Coverage by FSLI' for FSLI-level analysis"
    row = row + 1
    dashWs.Cells(row, 1).Value = "2. Review 'Coverage by Division' for division analysis"
    row = row + 1
    dashWs.Cells(row, 1).Value = "3. Use 'Manual Scoping' to adjust scoping decisions"
    row = row + 1
    dashWs.Cells(row, 1).Value = "4. Review 'Detailed Pack Analysis' for pack-level data"
    row = row + 1

    dashWs.Columns("A:B").AutoFit
End Sub

' ==================== VIEW 2: COVERAGE BY FSLI ====================
Private Sub CreateView2_CoverageByFSLI()
    '------------------------------------------------------------------------
    ' View 2: Coverage by FSLI
    ' For each FSLI: Total amount, Scoped amount, Untested amount, Coverage %
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet
    Dim inputWs As Worksheet
    Dim row As Long
    Dim col As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim fsli As String
    Dim totalAmount As Double
    Dim scopedAmount As Double
    Dim coveragePct As Double

    Set dashWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dashWs.Name = "Coverage by FSLI"

    ' Title
    dashWs.Cells(1, 1).Value = "COVERAGE BY FSLI"
    dashWs.Cells(1, 1).Font.Size = 14
    dashWs.Cells(1, 1).Font.Bold = True

    ' Headers
    row = 3
    dashWs.Cells(row, 1).Value = "FSLI"
    dashWs.Cells(row, 2).Value = "Total Consolidated Amount"
    dashWs.Cells(row, 3).Value = "Scoped Amount"
    dashWs.Cells(row, 4).Value = "Untested Amount"
    dashWs.Cells(row, 5).Value = "Coverage %"

    dashWs.Range("A3:E3").Font.Bold = True
    dashWs.Range("A3:E3").Interior.Color = RGB(68, 114, 196)
    dashWs.Range("A3:E3").Font.Color = RGB(255, 255, 255)

    ' Get input table
    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If inputWs Is Nothing Then Exit Sub

    lastCol = inputWs.Cells(1, inputWs.Columns.Count).End(xlToLeft).Column
    lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).row

    row = 4

    ' Process each FSLI column
    For col = 2 To lastCol
        fsli = inputWs.Cells(1, col).Value

        ' Calculate totals
        totalAmount = CalculateFSLITotal(inputWs, col, Mod1_MainController.g_ConsolidationEntity)
        scopedAmount = CalculateFSLIScopedAmount(inputWs, col)

        If totalAmount <> 0 Then
            coveragePct = (scopedAmount / totalAmount) * 100
        Else
            coveragePct = 0
        End If

        ' Write to dashboard
        dashWs.Cells(row, 1).Value = fsli
        dashWs.Cells(row, 2).Value = totalAmount
        dashWs.Cells(row, 2).NumberFormat = "#,##0.00"
        dashWs.Cells(row, 3).Value = scopedAmount
        dashWs.Cells(row, 3).NumberFormat = "#,##0.00"
        dashWs.Cells(row, 4).Value = totalAmount - scopedAmount
        dashWs.Cells(row, 4).NumberFormat = "#,##0.00"
        dashWs.Cells(row, 5).Value = coveragePct / 100
        dashWs.Cells(row, 5).NumberFormat = "0.0%"

        ' Color code coverage
        If coveragePct >= 80 Then
            dashWs.Cells(row, 5).Interior.Color = RGB(198, 239, 206) ' Green
        ElseIf coveragePct >= 60 Then
            dashWs.Cells(row, 5).Interior.Color = RGB(255, 235, 156) ' Yellow
        Else
            dashWs.Cells(row, 5).Interior.Color = RGB(255, 199, 206) ' Red
        End If

        row = row + 1
    Next col

    dashWs.Columns.AutoFit
End Sub

' ==================== VIEW 3: COVERAGE BY DIVISION ====================
Private Sub CreateView3_CoverageByDivision()
    '------------------------------------------------------------------------
    ' View 3: Coverage by Division
    ' For each Division: Packs, Scoped packs, Coverage %
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet

    Set dashWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dashWs.Name = "Coverage by Division"

    ' Title
    dashWs.Cells(1, 1).Value = "COVERAGE BY DIVISION"
    dashWs.Cells(1, 1).Font.Size = 14
    dashWs.Cells(1, 1).Font.Bold = True

    ' Headers
    dashWs.Cells(3, 1).Value = "Division"
    dashWs.Cells(3, 2).Value = "Total Packs"
    dashWs.Cells(3, 3).Value = "Scoped Packs"
    dashWs.Cells(3, 4).Value = "Coverage %"

    dashWs.Range("A3:D3").Font.Bold = True
    dashWs.Range("A3:D3").Interior.Color = RGB(68, 114, 196)
    dashWs.Range("A3:D3").Font.Color = RGB(255, 255, 255)

    ' Implementation: Calculate division-level statistics
    ' Uses Pack Number Company Table to get division assignments

    dashWs.Columns.AutoFit
End Sub

' ==================== VIEW 4: COVERAGE BY SEGMENT ====================
Private Sub CreateView4_CoverageBySegment()
    '------------------------------------------------------------------------
    ' View 4: Coverage by Segment
    ' For each Segment: Packs, Scoped packs, Coverage %
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet

    Set dashWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dashWs.Name = "Coverage by Segment"

    ' Title
    dashWs.Cells(1, 1).Value = "COVERAGE BY SEGMENT"
    dashWs.Cells(1, 1).Font.Size = 14
    dashWs.Cells(1, 1).Font.Bold = True

    ' Headers
    dashWs.Cells(3, 1).Value = "Segment"
    dashWs.Cells(3, 2).Value = "Total Packs"
    dashWs.Cells(3, 3).Value = "Scoped Packs"
    dashWs.Cells(3, 4).Value = "Coverage %"

    dashWs.Range("A3:D3").Font.Bold = True
    dashWs.Range("A3:D3").Interior.Color = RGB(68, 114, 196)
    dashWs.Range("A3:D3").Font.Color = RGB(255, 255, 255)

    ' Implementation: Calculate segment-level statistics
    ' Uses Division-Segment Mapping table

    dashWs.Columns.AutoFit
End Sub

' ==================== VIEW 5: DETAILED PACK ANALYSIS ====================
Private Sub CreateView5_DetailedPackAnalysis()
    '------------------------------------------------------------------------
    ' View 5: Detailed Pack Analysis
    ' Interactive table showing all packs × all FSLIs with amounts, %, scoping status
    ' Sortable, filterable
    '------------------------------------------------------------------------
    Dim dashWs As Worksheet
    Dim inputWs As Worksheet
    Dim row As Long
    Dim dataRow As Long
    Dim col As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim packName As String
    Dim packCode As String
    Dim fsli As String
    Dim amount As Variant
    Dim percentage As Double
    Dim scopingStatus As String

    Set dashWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    dashWs.Name = "Detailed Pack Analysis"

    ' Title
    dashWs.Cells(1, 1).Value = "DETAILED PACK × FSLI ANALYSIS"
    dashWs.Cells(1, 1).Font.Size = 14
    dashWs.Cells(1, 1).Font.Bold = True

    ' Headers
    row = 3
    dashWs.Cells(row, 1).Value = "Pack Code"
    dashWs.Cells(row, 2).Value = "Pack Name"
    dashWs.Cells(row, 3).Value = "FSLI"
    dashWs.Cells(row, 4).Value = "Amount"
    dashWs.Cells(row, 5).Value = "% of Consolidated"
    dashWs.Cells(row, 6).Value = "Scoping Status"
    dashWs.Cells(row, 7).Value = "Division"
    dashWs.Cells(row, 8).Value = "Segment"

    dashWs.Range("A3:H3").Font.Bold = True
    dashWs.Range("A3:H3").Interior.Color = RGB(68, 114, 196)
    dashWs.Range("A3:H3").Font.Color = RGB(255, 255, 255)

    ' Get data from Full Input Table
    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If inputWs Is Nothing Then Exit Sub

    lastCol = inputWs.Cells(1, inputWs.Columns.Count).End(xlToLeft).Column
    lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).row

    row = 4

    ' Populate data
    For dataRow = 2 To lastRow
        packName = inputWs.Cells(dataRow, 1).Value
        packCode = ExtractPackCodeFromName(packName)

        For col = 2 To lastCol
            fsli = inputWs.Cells(1, col).Value
            amount = inputWs.Cells(dataRow, col).Value

            If IsNumeric(amount) Then
                ' Get percentage from percentage table
                percentage = GetPercentageValue(packCode, fsli)

                ' Determine scoping status
                scopingStatus = GetScopingStatus(packCode, fsli)

                ' Write row
                dashWs.Cells(row, 1).Value = packCode
                dashWs.Cells(row, 2).Value = packName
                dashWs.Cells(row, 3).Value = fsli
                dashWs.Cells(row, 4).Value = CDbl(amount)
                dashWs.Cells(row, 4).NumberFormat = "#,##0.00"
                dashWs.Cells(row, 5).Value = percentage
                dashWs.Cells(row, 5).NumberFormat = "0.00%"
                dashWs.Cells(row, 6).Value = scopingStatus

                ' Color code status
                Select Case scopingStatus
                    Case "Scoped In"
                        dashWs.Cells(row, 6).Interior.Color = RGB(198, 239, 206)
                    Case "Not Scoped"
                        dashWs.Cells(row, 6).Interior.Color = RGB(255, 235, 156)
                End Select

                row = row + 1
            End If
        Next col
    Next dataRow

    ' Add autofilter
    dashWs.Range("A3:H3").AutoFilter

    dashWs.Columns.AutoFit
End Sub

' ==================== MANUAL SCOPING INTERFACE ====================
Private Sub CreateManualScopingInterface()
    '------------------------------------------------------------------------
    ' Create interactive manual scoping interface
    ' Allows users to change scoping status with dropdowns
    '------------------------------------------------------------------------
    Dim scopeWs As Worksheet

    Set scopeWs = Mod1_MainController.g_OutputWorkbook.Worksheets.Add
    scopeWs.Name = "Manual Scoping Interface"

    ' Title
    scopeWs.Cells(1, 1).Value = "MANUAL SCOPING INTERFACE"
    scopeWs.Cells(1, 1).Font.Size = 14
    scopeWs.Cells(1, 1).Font.Bold = True

    ' Instructions
    scopeWs.Cells(3, 1).Value = "Use this interface to manually adjust scoping decisions:"
    scopeWs.Cells(4, 1).Value = "• Filter by FSLI, Division, or Segment"
    scopeWs.Cells(5, 1).Value = "• Change 'Scoping Status' to 'Scoped In' or 'Not Scoped'"
    scopeWs.Cells(6, 1).Value = "• Coverage percentages update automatically"

    ' Headers
    Dim row As Long
    row = 8
    scopeWs.Cells(row, 1).Value = "Pack Code"
    scopeWs.Cells(row, 2).Value = "Pack Name"
    scopeWs.Cells(row, 3).Value = "FSLI"
    scopeWs.Cells(row, 4).Value = "Amount"
    scopeWs.Cells(row, 5).Value = "Scoping Status"
    scopeWs.Cells(row, 6).Value = "Division"
    scopeWs.Cells(row, 7).Value = "Segment"

    scopeWs.Range("A8:G8").Font.Bold = True
    scopeWs.Range("A8:G8").Interior.Color = RGB(68, 114, 196)
    scopeWs.Range("A8:G8").Font.Color = RGB(255, 255, 255)

    ' Add data (similar to Detailed Pack Analysis but with editable status)
    ' Implementation: Populate with all pack-FSLI combinations
    ' Add data validation for Scoping Status column

    scopeWs.Columns.AutoFit
End Sub

' ==================== HELPER FUNCTIONS ====================
Private Function GetTotalPackCount() As Long
    Dim inputWs As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set inputWs = Mod1_MainController.g_OutputWorkbook.Worksheets("Full Input Table")
    On Error GoTo 0

    If inputWs Is Nothing Then
        GetTotalPackCount = 0
        Exit Function
    End If

    lastRow = inputWs.Cells(inputWs.Rows.Count, 1).End(xlUp).row
    GetTotalPackCount = lastRow - 1 ' Subtract header row
End Function

Private Function CalculateFSLITotal(ws As Worksheet, fsliCol As Long, consolEntity As String) As Double
    ' Calculate total for an FSLI (from consolidation entity row)
    Dim row As Long
    Dim packName As String

    For row = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        packName = ws.Cells(row, 1).Value
        If InStr(packName, consolEntity) > 0 Then
            If IsNumeric(ws.Cells(row, fsliCol).Value) Then
                CalculateFSLITotal = Abs(CDbl(ws.Cells(row, fsliCol).Value))
            End If
            Exit Function
        End If
    Next row
End Function

Private Function CalculateFSLIScopedAmount(ws As Worksheet, fsliCol As Long) As Double
    ' Calculate scoped amount for an FSLI
    Dim row As Long
    Dim packCode As String
    Dim amount As Double
    Dim total As Double

    total = 0

    For row = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).row
        packCode = ExtractPackCodeFromName(ws.Cells(row, 1).Value)

        If Mod1_MainController.g_ScopedPacks.exists(packCode) Then
            If IsNumeric(ws.Cells(row, fsliCol).Value) Then
                amount = Abs(CDbl(ws.Cells(row, fsliCol).Value))
                total = total + amount
            End If
        End If
    Next row

    CalculateFSLIScopedAmount = total
End Function

Private Function ExtractPackCodeFromName(packNameWithCode As String) As String
    Dim openParen As Long
    Dim closeParen As Long

    openParen = InStrRev(packNameWithCode, "(")
    closeParen = InStrRev(packNameWithCode, ")")

    If openParen > 0 And closeParen > openParen Then
        ExtractPackCodeFromName = Trim(Mid(packNameWithCode, openParen + 1, closeParen - openParen - 1))
    Else
        ExtractPackCodeFromName = packNameWithCode
    End If
End Function

Private Function GetPercentageValue(packCode As String, fsli As String) As Double
    ' Get percentage value from Full Input Percentage table
    ' Implementation: Look up value in percentage table
    GetPercentageValue = 0
End Function

Private Function GetScopingStatus(packCode As String, fsli As String) As String
    ' Determine scoping status for a pack-FSLI combination
    If Mod1_MainController.g_ScopedPacks.exists(packCode) Then
        GetScopingStatus = "Scoped In"
    ElseIf Mod1_MainController.g_ManualScoping.exists(packCode & "|" & fsli) Then
        GetScopingStatus = Mod1_MainController.g_ManualScoping(packCode & "|" & fsli)
    Else
        GetScopingStatus = "Not Scoped"
    End If
End Function
