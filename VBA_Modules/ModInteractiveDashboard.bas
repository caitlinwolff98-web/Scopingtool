Attribute VB_Name = "ModInteractiveDashboard"
Option Explicit

' ============================================================================
' MODULE: ModInteractiveDashboard
' PURPOSE: Create interactive Excel dashboard with slicers and pivot tables
' DESCRIPTION: Adds Excel-based interactivity so the workbook can be used
'              without PowerBI for scoping analysis
' ============================================================================

' Create the interactive Excel dashboard
Public Sub CreateInteractiveDashboard()
    On Error GoTo ErrorHandler
    
    Dim dashboardWs As Worksheet
    
    ' Check if dashboard sheet already exists
    On Error Resume Next
    Set dashboardWs = g_OutputWorkbook.Worksheets("Interactive Dashboard")
    On Error GoTo ErrorHandler
    
    If dashboardWs Is Nothing Then
        Set dashboardWs = g_OutputWorkbook.Worksheets.Add
        dashboardWs.Name = "Interactive Dashboard"
    Else
        dashboardWs.Cells.Clear
    End If
    
    ' Create dashboard layout
    CreateDashboardLayout dashboardWs
    
    ' Create pivot tables for analysis
    CreateScopingPivotTable dashboardWs
    
    ' Create summary charts
    CreateSummaryCharts dashboardWs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating interactive dashboard: " & Err.Description, vbCritical
End Sub

' Create dashboard layout with instructions and key metrics
Private Sub CreateDashboardLayout(ws As Worksheet)
    On Error Resume Next
    
    Dim row As Long
    
    row = 1
    With ws
        ' Title
        .Cells(row, 1).Value = "BIDVEST SCOPING TOOL - INTERACTIVE DASHBOARD"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 16
        .Cells(row, 1).Font.Color = RGB(68, 114, 196)
        row = row + 2
        
        ' Instructions
        .Cells(row, 1).Value = "Instructions:"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "1. Use the Scoping Summary sheet to review pack-level recommendations"
        row = row + 1
        .Cells(row, 1).Value = "2. Use the Threshold Configuration sheet (if applicable) to see automatic scoping rules"
        row = row + 1
        .Cells(row, 1).Value = "3. Review the data tables (Full Input, Console, etc.) for detailed analysis"
        row = row + 1
        .Cells(row, 1).Value = "4. Use this dashboard for high-level summaries and pivot analysis"
        row = row + 2
        
        ' Key Metrics Section
        .Cells(row, 1).Value = "KEY METRICS"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 12
        row = row + 1
        
        ' Calculate metrics from Scoping Summary if it exists
        Dim summaryWs As Worksheet
        Set summaryWs = g_OutputWorkbook.Worksheets("Scoping Summary")
        
        If Not summaryWs Is Nothing Then
            Dim totalPacks As Long
            Dim scopedPacks As Long
            Dim lastRow As Long
            Dim i As Long
            
            lastRow = summaryWs.Cells(summaryWs.Rows.Count, 1).End(xlUp).Row
            
            For i = 4 To lastRow ' Start after headers
                If summaryWs.Cells(i, 1).Value <> "" Then
                    totalPacks = totalPacks + 1
                    If InStr(1, summaryWs.Cells(i, 3).Value, "Yes", vbTextCompare) > 0 Then
                        scopedPacks = scopedPacks + 1
                    End If
                End If
            Next i
            
            .Cells(row, 1).Value = "Total Packs:"
            .Cells(row, 2).Value = totalPacks
            .Cells(row, 2).Font.Bold = True
            row = row + 1
            
            .Cells(row, 1).Value = "Scoped In:"
            .Cells(row, 2).Value = scopedPacks
            .Cells(row, 2).Font.Bold = True
            .Cells(row, 2).Interior.Color = RGB(198, 239, 206)
            row = row + 1
            
            .Cells(row, 1).Value = "Pending Review:"
            .Cells(row, 2).Value = totalPacks - scopedPacks
            .Cells(row, 2).Font.Bold = True
            .Cells(row, 2).Interior.Color = RGB(255, 235, 156)
            row = row + 1
            
            If totalPacks > 0 Then
                .Cells(row, 1).Value = "Coverage:"
                .Cells(row, 2).Value = Format(scopedPacks / totalPacks, "0.0%")
                .Cells(row, 2).Font.Bold = True
            End If
        End If
        
        ' Auto-fit columns
        .Columns("A:B").AutoFit
    End With
    
    On Error GoTo 0
End Sub

' Create pivot table for scoping analysis
Private Sub CreateScopingPivotTable(ws As Worksheet)
    On Error Resume Next
    
    Dim summaryWs As Worksheet
    Dim pvtCache As PivotCache
    Dim pvtTable As PivotTable
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sourceRange As Range
    
    ' Get Scoping Summary sheet
    Set summaryWs = g_OutputWorkbook.Worksheets("Scoping Summary")
    If summaryWs Is Nothing Then Exit Sub
    
    ' Find the table range
    lastRow = summaryWs.Cells(summaryWs.Rows.Count, 1).End(xlUp).Row
    lastCol = 4 ' Pack Code, Pack Name, Scoped In, Suggested for Scope
    
    If lastRow < 4 Then Exit Sub ' Not enough data
    
    ' Set source range (the table in Scoping Summary)
    Set sourceRange = summaryWs.Range(summaryWs.Cells(3, 1), summaryWs.Cells(lastRow, lastCol))
    
    ' Create pivot cache
    Set pvtCache = g_OutputWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sourceRange)
    
    ' Create pivot table
    Set pvtTable = pvtCache.CreatePivotTable( _
        TableDestination:=ws.Range("A15"), _
        TableName:="ScopingAnalysisPivot")
    
    ' Configure pivot table
    With pvtTable
        ' Add Scoped In to Row area
        .PivotFields("Scoped In").Orientation = xlRowField
        
        ' Add Pack Name to Row area
        .PivotFields("Pack Name").Orientation = xlRowField
        
        ' Add count to Values area
        .AddDataField .PivotFields("Pack Code"), "Count of Packs", xlCount
        
        ' Format
        .TableStyle2 = "PivotStyleMedium9"
    End With
    
    On Error GoTo 0
End Sub

' Create summary charts
Private Sub CreateSummaryCharts(ws As Worksheet)
    On Error Resume Next
    
    Dim chartObj As ChartObject
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    Dim scopedCount As Long
    Dim reviewCount As Long
    Dim i As Long
    
    ' Get counts from Scoping Summary
    Set summaryWs = g_OutputWorkbook.Worksheets("Scoping Summary")
    If summaryWs Is Nothing Then Exit Sub
    
    lastRow = summaryWs.Cells(summaryWs.Rows.Count, 1).End(xlUp).Row
    
    For i = 4 To lastRow
        If summaryWs.Cells(i, 1).Value <> "" Then
            If InStr(1, summaryWs.Cells(i, 3).Value, "Yes", vbTextCompare) > 0 Then
                scopedCount = scopedCount + 1
            Else
                reviewCount = reviewCount + 1
            End If
        End If
    Next i
    
    ' Create temporary data for chart
    ws.Cells(20, 5).Value = "Status"
    ws.Cells(20, 6).Value = "Count"
    ws.Cells(21, 5).Value = "Scoped In"
    ws.Cells(21, 6).Value = scopedCount
    ws.Cells(22, 5).Value = "Pending Review"
    ws.Cells(22, 6).Value = reviewCount
    
    ' Create pie chart
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Cells(15, 5).Left, _
        Top:=ws.Cells(15, 5).Top, _
        Width:=300, _
        Height:=200)
    
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("E20:F22")
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Scoping Status"
        .ApplyLayout 1
    End With
    
    On Error GoTo 0
End Sub

' Add data validation and dropdowns for interactive filtering
Public Sub AddInteractiveFilters()
    On Error Resume Next
    
    Dim summaryWs As Worksheet
    Dim lastRow As Long
    
    Set summaryWs = g_OutputWorkbook.Worksheets("Scoping Summary")
    If summaryWs Is Nothing Then Exit Sub
    
    lastRow = summaryWs.Cells(summaryWs.Rows.Count, 1).End(xlUp).Row
    
    ' Add filter buttons to the table headers
    If summaryWs.AutoFilterMode Then
        summaryWs.AutoFilterMode = False
    End If
    
    summaryWs.Range("A3:D" & lastRow).AutoFilter
    
    On Error GoTo 0
End Sub

' Create a simple scoping calculator
Public Sub CreateScopingCalculator()
    On Error Resume Next
    
    Dim calcWs As Worksheet
    Dim row As Long
    
    ' Check if sheet exists
    On Error Resume Next
    Set calcWs = g_OutputWorkbook.Worksheets("Scoping Calculator")
    On Error GoTo 0
    
    If calcWs Is Nothing Then
        Set calcWs = g_OutputWorkbook.Worksheets.Add
        calcWs.Name = "Scoping Calculator"
    Else
        calcWs.Cells.Clear
    End If
    
    row = 1
    With calcWs
        .Cells(row, 1).Value = "SCOPING CALCULATOR"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        row = row + 2
        
        .Cells(row, 1).Value = "Use this calculator to estimate scoping coverage:"
        row = row + 2
        
        .Cells(row, 1).Value = "Total Number of Packs:"
        .Cells(row, 2).Interior.Color = RGB(255, 242, 204)
        row = row + 1
        
        .Cells(row, 1).Value = "Number of Packs to Scope:"
        .Cells(row, 2).Interior.Color = RGB(255, 242, 204)
        row = row + 1
        
        .Cells(row, 1).Value = "Coverage Percentage:"
        .Cells(row, 2).Formula = "=IF(B4>0,B5/B4,0)"
        .Cells(row, 2).NumberFormat = "0.0%"
        .Cells(row, 2).Font.Bold = True
        row = row + 2
        
        .Cells(row, 1).Value = "Target Coverage:"
        .Cells(row, 2).Value = 0.8
        .Cells(row, 2).NumberFormat = "0.0%"
        .Cells(row, 2).Interior.Color = RGB(255, 242, 204)
        row = row + 1
        
        .Cells(row, 1).Value = "Packs Needed for Target:"
        .Cells(row, 2).Formula = "=B4*B9"
        .Cells(row, 2).NumberFormat = "0"
        .Cells(row, 2).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "Additional Packs Needed:"
        .Cells(row, 2).Formula = "=MAX(0,B10-B5)"
        .Cells(row, 2).NumberFormat = "0"
        .Cells(row, 2).Font.Bold = True
        
        .Columns("A:B").AutoFit
    End With
    
    On Error GoTo 0
End Sub
