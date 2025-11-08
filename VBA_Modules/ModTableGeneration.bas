Attribute VB_Name = "ModTableGeneration"
Option Explicit

' ============================================================================
' MODULE: ModTableGeneration
' PURPOSE: Generate supporting tables (FSLi Key, Pack Number, Percentages)
' DESCRIPTION: Creates additional tables required for Power BI integration
' ============================================================================

' Create FSLi Key Table with all FSLi entries
Public Sub CreateFSLiKeyTable()
    On Error GoTo ErrorHandler
    
    Dim outputWs As Worksheet
    Dim fsliCollection As Collection
    Dim fsliName As String
    Dim row As Long
    Dim i As Long
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "FSLi Key Table"
    
    ' Set up headers
    outputWs.Cells(1, 1).Value = "FSLi"
    outputWs.Cells(1, 2).Value = "FSLi Input"
    outputWs.Cells(1, 3).Value = "FSLi Input Percentage"
    outputWs.Cells(1, 4).Value = "FSLi Journal"
    outputWs.Cells(1, 5).Value = "FSLi Journal Percentage"
    outputWs.Cells(1, 6).Value = "FSLi Console"
    outputWs.Cells(1, 7).Value = "FSLi Console Percentage"
    outputWs.Cells(1, 8).Value = "FSLi Discontinued"
    outputWs.Cells(1, 9).Value = "FSLi Discontinued Percentage"
    
    ' Get unique FSLi names from all tables
    Set fsliCollection = CollectAllFSLiNames()
    
    ' Populate FSLi names
    row = 2
    For i = 1 To fsliCollection.Count
        fsliName = fsliCollection(i)
        outputWs.Cells(row, 1).Value = fsliName
        
        ' Link to other tables using VLOOKUP formulas
        ' FSLi Input
        outputWs.Cells(row, 2).Formula = "=IFERROR(VLOOKUP(A" & row & ",'Full Input Table'!A:B,2,FALSE),0)"
        
        ' FSLi Input Percentage
        outputWs.Cells(row, 3).Formula = "=IFERROR(VLOOKUP(A" & row & ",'Full Input Percentage'!A:B,2,FALSE),0)"
        
        ' Similar formulas for other columns...
        
        row = row + 1
    Next i
    
    ' Format table
    FormatAsTable outputWs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating FSLi Key Table: " & Err.Description, vbCritical
End Sub

' Collect all unique FSLi names from all tables
Private Function CollectAllFSLiNames() As Collection
    Dim fsliNames As Object ' Dictionary
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim col As Long
    Dim fsliName As String
    Dim resultCollection As New Collection
    
    Set fsliNames = CreateObject("Scripting.Dictionary")
    
    ' Check each generated table
    For Each ws In g_OutputWorkbook.Worksheets
        If ws.Name Like "*Table" And ws.Name <> "FSLi Key Table" And _
           ws.Name <> "Pack Number Company Table" Then
            
            ' Get FSLi names from row 1 (starting from column 2)
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            For col = 2 To lastCol
                fsliName = Trim(ws.Cells(1, col).Value)
                If fsliName <> "" Then
                    ' Remove metadata tags like (Total) or (Subtotal)
                    fsliName = RemoveMetadataTags(fsliName)
                    
                    If Not fsliNames.Exists(fsliName) Then
                        fsliNames.Add fsliName, True
                        resultCollection.Add fsliName
                    End If
                End If
            Next col
        End If
    Next ws
    
    Set CollectAllFSLiNames = resultCollection
End Function

' Remove metadata tags from FSLi names
Private Function RemoveMetadataTags(fsliName As String) As String
    Dim cleanName As String
    cleanName = fsliName
    
    ' Remove common tags
    cleanName = Replace(cleanName, " (Total)", "")
    cleanName = Replace(cleanName, " (Subtotal)", "")
    cleanName = Trim(cleanName)
    
    RemoveMetadataTags = cleanName
End Function

' Create Pack Number Company Table with divisions
Public Sub CreatePackNumberCompanyTable()
    On Error GoTo ErrorHandler
    
    Dim outputWs As Worksheet
    Dim segmentTabs As Collection
    Dim discontinuedTab As Worksheet
    Dim packDict As Object ' Dictionary to avoid duplicates
    Dim tabInfo As Object
    Dim ws As Worksheet
    Dim row As Long
    Dim i As Long
    Dim packName As String
    Dim packCode As String
    Dim divisionName As String
    Dim col As Long
    Dim lastCol As Long
    
    Set packDict = CreateObject("Scripting.Dictionary")
    
    ' Get segment tabs
    Set segmentTabs = GetTabsForCategory(CAT_SEGMENT)
    
    ' Process each segment tab
    For i = 1 To segmentTabs.Count
        Set tabInfo = segmentTabs(i)
        Set ws = g_SourceWorkbook.Worksheets(tabInfo("TabName"))
        
        ' Get division name
        divisionName = tabInfo("DivisionName")
        If divisionName = "" Then
            divisionName = PromptForDivisionName(tabInfo("TabName"))
            tabInfo("DivisionName") = divisionName
        End If
        
        ' Extract pack names from row 7 and codes from row 8
        lastCol = ws.Cells(7, ws.Columns.Count).End(xlToLeft).Column
        
        For col = 1 To lastCol
            packName = Trim(ws.Cells(7, col).Value)
            packCode = Trim(ws.Cells(8, col).Value)
            
            If packName <> "" And packCode <> "" Then
                ' Create unique key
                Dim packKey As String
                packKey = packCode
                
                If Not packDict.Exists(packKey) Then
                    Dim packInfo As Object
                    Set packInfo = CreateObject("Scripting.Dictionary")
                    packInfo("Name") = packName
                    packInfo("Code") = packCode
                    packInfo("Division") = divisionName
                    packDict.Add packKey, packInfo
                End If
            End If
        Next col
    Next i
    
    ' Process discontinued tab if it exists
    Set discontinuedTab = GetTabByCategory(CAT_DISCONTINUED)
    If Not discontinuedTab Is Nothing Then
        lastCol = discontinuedTab.Cells(7, discontinuedTab.Columns.Count).End(xlToLeft).Column
        
        For col = 1 To lastCol
            packName = Trim(discontinuedTab.Cells(7, col).Value)
            packCode = Trim(discontinuedTab.Cells(8, col).Value)
            
            If packName <> "" And packCode <> "" Then
                packKey = packCode
                
                If Not packDict.Exists(packKey) Then
                    Set packInfo = CreateObject("Scripting.Dictionary")
                    packInfo("Name") = packName
                    packInfo("Code") = packCode
                    packInfo("Division") = "Discontinued"
                    packDict.Add packKey, packInfo
                End If
            End If
        Next col
    End If
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Pack Number Company Table"
    
    ' Set up headers
    outputWs.Cells(1, 1).Value = "Pack Name"
    outputWs.Cells(1, 2).Value = "Pack Code"
    outputWs.Cells(1, 3).Value = "Division"
    
    ' Write data
    row = 2
    Dim key As Variant
    For Each key In packDict.Keys
        Set packInfo = packDict(key)
        outputWs.Cells(row, 1).Value = packInfo("Name")
        outputWs.Cells(row, 2).Value = packInfo("Code")
        outputWs.Cells(row, 3).Value = packInfo("Division")
        row = row + 1
    Next key
    
    ' Format table
    FormatAsTable outputWs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Pack Number Company Table: " & Err.Description, vbCritical
End Sub

' Prompt user for division name
Private Function PromptForDivisionName(tabName As String) As String
    Dim divName As String
    
    divName = InputBox( _
        "Please enter the division name for the segment tab:" & vbCrLf & vbCrLf & _
        "Tab: " & tabName & vbCrLf & vbCrLf & _
        "Example: If this is 'TGK UK' tab, enter 'UK'", _
        "Enter Division Name", _
        "")
    
    PromptForDivisionName = Trim(divName)
End Function

' Create percentage tables for all main tables
Public Sub CreatePercentageTables()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim tableName As String
    
    ' Process each main table
    For Each ws In g_OutputWorkbook.Worksheets
        tableName = ws.Name
        
        ' Only process main data tables
        If tableName = "Full Input Table" Or _
           tableName = "Journals Table" Or _
           tableName = "Full Console Table" Or _
           tableName = "Discontinued Table" Then
            
            CreatePercentageTable ws
        End If
    Next ws
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating percentage tables: " & Err.Description, vbCritical
End Sub

' Create percentage table for a specific data table
Private Sub CreatePercentageTable(sourceWs As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim outputWs As Worksheet
    Dim percentTableName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim row As Long
    Dim col As Long
    Dim cellValue As Variant
    Dim columnTotal As Double
    Dim percentValue As Double
    
    ' Create percentage table name
    percentTableName = Replace(sourceWs.Name, "Table", "Percentage")
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = percentTableName
    
    ' Get dimensions
    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).row
    lastCol = sourceWs.Cells(1, sourceWs.Columns.Count).End(xlToLeft).Column
    
    ' Copy headers
    sourceWs.Rows(1).Copy outputWs.Rows(1)
    
    ' Calculate percentages for each column
    For col = 2 To lastCol ' Start from column 2 (skip Pack names)
        ' Calculate column total
        columnTotal = 0
        For row = 2 To lastRow
            cellValue = sourceWs.Cells(row, col).Value
            If IsNumeric(cellValue) Then
                columnTotal = columnTotal + Abs(cellValue)
            End If
        Next row
        
        ' Calculate percentage for each cell
        For row = 2 To lastRow
            cellValue = sourceWs.Cells(row, col).Value
            
            If IsNumeric(cellValue) And columnTotal <> 0 Then
                percentValue = (Abs(cellValue) / columnTotal) * 100
                outputWs.Cells(row, col).Value = percentValue
                outputWs.Cells(row, col).NumberFormat = "0.00%"
            Else
                outputWs.Cells(row, col).Value = 0
                outputWs.Cells(row, col).NumberFormat = "0.00%"
            End If
        Next row
    Next col
    
    ' Copy pack names
    For row = 2 To lastRow
        outputWs.Cells(row, 1).Value = sourceWs.Cells(row, 1).Value
    Next row
    
    ' Format table
    FormatAsTable outputWs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating percentage table for " & sourceWs.Name & ": " & Err.Description, vbCritical
End Sub

' Format worksheet as table (shared utility)
Private Sub FormatAsTable(ws As Worksheet)
    On Error Resume Next
    
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tableRange As Range
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow > 1 And lastCol > 1 Then
        Set tableRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        
        ' Format headers
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).Interior.Color = RGB(68, 114, 196)
        ws.Rows(1).Font.Color = RGB(255, 255, 255)
        
        ' Auto-fit columns
        ws.Columns.AutoFit
        
        ' Add borders
        tableRange.Borders.LineStyle = xlContinuous
        
        ' Freeze top row
        ws.Rows(2).Select
        ActiveWindow.FreezePanes = True
        ws.Range("A1").Select
    End If
    
    On Error GoTo 0
End Sub
