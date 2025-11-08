Attribute VB_Name = "ModDataProcessing"
Option Explicit

' ============================================================================
' MODULE: ModDataProcessing
' PURPOSE: Process consolidation data and analyze structure
' DESCRIPTION: Handles unmerging cells, detecting columns, analyzing FSLi
'              hierarchies, and extracting entity information
' ============================================================================

' Structure to hold column information
Public Type ColumnInfo
    ColumnIndex As Long
    ColumnType As String ' "Original/Entity" or "Consolidation/Consolidation"
    PackName As String
    PackCode As String
End Type

' Structure to hold FSLi information
Public Type FSLiInfo
    FSLiName As String
    RowIndex As Long
    IsTotal As Boolean
    IsSubtotal As Boolean
    SubtotalOf As String
    Level As Long
    StatementType As String ' "Income Statement", "Balance Sheet", "Notes"
End Type

' Main processing orchestrator
Public Sub ProcessConsolidationData()
    On Error GoTo ErrorHandler
    
    Dim inputTab As Worksheet
    Dim discontinuedTab As Worksheet
    Dim journalsTab As Worksheet
    Dim consoleTab As Worksheet
    
    ' Get required tabs
    Set inputTab = GetTabByCategory(CAT_INPUT_CONTINUING)
    
    If inputTab Is Nothing Then
        MsgBox "Could not find Input Continuing tab. Process cannot continue.", vbCritical
        Exit Sub
    End If
    
    ' Process Input Continuing tab
    Application.StatusBar = "Processing Input Continuing tab..."
    ProcessInputTab inputTab
    
    ' Process other tabs if they exist
    Set discontinuedTab = GetTabByCategory(CAT_DISCONTINUED)
    If Not discontinuedTab Is Nothing Then
        Application.StatusBar = "Processing Discontinued tab..."
        ProcessDiscontinuedTab discontinuedTab
    End If
    
    Set journalsTab = GetTabByCategory(CAT_JOURNALS_CONTINUING)
    If Not journalsTab Is Nothing Then
        Application.StatusBar = "Processing Journals tab..."
        ProcessJournalsTab journalsTab
    End If
    
    Set consoleTab = GetTabByCategory(CAT_CONSOLE_CONTINUING)
    If Not consoleTab Is Nothing Then
        Application.StatusBar = "Processing Console tab..."
        ProcessConsoleTab consoleTab
    End If
    
    ' Create supporting tables
    Application.StatusBar = "Creating FSLi Key Table..."
    CreateFSLiKeyTable
    
    Application.StatusBar = "Creating Pack Number Company Table..."
    CreatePackNumberCompanyTable
    
    Application.StatusBar = "Creating Percentage Tables..."
    CreatePercentageTables
    
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error in data processing: " & Err.Description, vbCritical
End Sub

' Get worksheet by category
Private Function GetTabByCategory(categoryName As String) As Worksheet
    On Error Resume Next
    Dim tabInfo As Object
    
    If g_TabCategories.Exists(categoryName) Then
        If g_TabCategories(categoryName).Count > 0 Then
            Set tabInfo = g_TabCategories(categoryName)(1)
            Set GetTabByCategory = g_SourceWorkbook.Worksheets(tabInfo("TabName"))
        End If
    End If
    
    On Error GoTo 0
End Function

' Process Input Continuing tab
Private Sub ProcessInputTab(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim lastCol As Long
    Dim lastRow As Long
    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String
    
    ' Step 1: Unmerge all cells
    ws.Cells.UnMerge
    
    ' Step 2: Detect columns and get user selection
    Set columns = DetectColumns(ws)
    selectedColumnType = PromptColumnSelection(columns)
    
    If selectedColumnType = "" Then
        MsgBox "No column type selected. Skipping Input tab.", vbExclamation
        Exit Sub
    End If
    
    ' Step 3: Analyze FSLi structure
    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    
    ' Step 4: Create Full Input Table
    CreateFullInputTable ws, columns, fsliList, selectedColumnType
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing Input tab: " & Err.Description, vbCritical
End Sub

' Detect columns in row 6
Private Function DetectColumns(ws As Worksheet) As Collection
    On Error GoTo ErrorHandler
    
    Dim columns As New Collection
    Dim col As Long
    Dim lastCol As Long
    Dim cellValue As String
    Dim colInfo As ColumnInfo
    Dim packName As String
    Dim packCode As String
    
    ' Find last column with data in row 6
    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column
    
    ' Analyze row 6 for column types
    For col = 1 To lastCol
        cellValue = Trim(ws.Cells(6, col).Value)
        
        If cellValue <> "" Then
            colInfo.ColumnIndex = col
            
            ' Determine column type
            If InStr(1, cellValue, "original", vbTextCompare) > 0 And _
               InStr(1, cellValue, "entity currency", vbTextCompare) > 0 Then
                colInfo.ColumnType = "Original/Entity"
            ElseIf InStr(1, cellValue, "consolidation", vbTextCompare) > 0 And _
                   InStr(1, cellValue, "consolidation currency", vbTextCompare) > 0 Then
                colInfo.ColumnType = "Consolidation/Consolidation"
            Else
                colInfo.ColumnType = "Other"
            End If
            
            ' Get pack name from row 7
            packName = ""
            If ws.Cells(7, col).Value <> "" Then
                packName = Trim(ws.Cells(7, col).Value)
            End If
            colInfo.PackName = packName
            
            ' Get pack code from row 8
            packCode = ""
            If ws.Cells(8, col).Value <> "" Then
                packCode = Trim(ws.Cells(8, col).Value)
            End If
            colInfo.PackCode = packCode
            
            columns.Add colInfo
        End If
    Next col
    
    Set DetectColumns = columns
    Exit Function
    
ErrorHandler:
    MsgBox "Error detecting columns: " & Err.Description, vbCritical
    Set DetectColumns = New Collection
End Function

' Prompt user to select column type
Private Function PromptColumnSelection(columns As Collection) As String
    Dim originalCount As Long
    Dim consolidationCount As Long
    Dim i As Long
    Dim colInfo As ColumnInfo
    Dim msg As String
    Dim response As VbMsgBoxResult
    
    ' Count column types
    For i = 1 To columns.Count
        colInfo = columns(i)
        If colInfo.ColumnType = "Original/Entity" Then
            originalCount = originalCount + 1
        ElseIf colInfo.ColumnType = "Consolidation/Consolidation" Then
            consolidationCount = consolidationCount + 1
        End If
    Next i
    
    ' Build message
    msg = "Column types detected in row 6:" & vbCrLf & vbCrLf
    
    If originalCount > 0 Then
        msg = msg & "- Original/Entity Currency: " & originalCount & " columns" & vbCrLf
    End If
    
    If consolidationCount > 0 Then
        msg = msg & "- Consolidation/Consolidation Currency: " & consolidationCount & " columns" & vbCrLf
    End If
    
    msg = msg & vbCrLf & "Which columns do you want to use?" & vbCrLf & vbCrLf
    msg = msg & "Click YES for Consolidation/Consolidation Currency (recommended)" & vbCrLf
    msg = msg & "Click NO for Original/Entity Currency"
    
    response = MsgBox(msg, vbYesNoCancel + vbQuestion, "Select Column Type")
    
    If response = vbYes Then
        PromptColumnSelection = "Consolidation/Consolidation"
    ElseIf response = vbNo Then
        PromptColumnSelection = "Original/Entity"
    Else
        PromptColumnSelection = ""
    End If
End Function

' Analyze FSLi structure
Private Function AnalyzeFSLiStructure(ws As Worksheet, columnType As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim fsliList As New Collection
    Dim row As Long
    Dim lastRow As Long
    Dim fsliName As String
    Dim fsliInfo As FSLiInfo
    Dim currentStatement As String
    Dim notesStartRow As Long
    
    currentStatement = ""
    notesStartRow = 0
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    
    ' Start from row 9 (after headers)
    For row = 9 To lastRow
        fsliName = Trim(ws.Cells(row, 2).Value)
        
        ' Check if this is the Notes section
        If UCase(fsliName) = "NOTES" Then
            notesStartRow = row
            Exit For
        End If
        
        ' Skip empty rows
        If fsliName = "" Then
            If Not IsRowEmpty(ws, row) Then
                ' Row has data but no FSLi name - might be a continuation
            End If
            GoTo NextRow
        End If
        
        ' Detect statement type
        If InStr(1, fsliName, "income statement", vbTextCompare) > 0 Then
            currentStatement = "Income Statement"
        ElseIf InStr(1, fsliName, "balance sheet", vbTextCompare) > 0 Then
            currentStatement = "Balance Sheet"
        End If
        
        ' Create FSLi info
        fsliInfo.FSLiName = fsliName
        fsliInfo.RowIndex = row
        fsliInfo.StatementType = currentStatement
        
        ' Detect if it's a total or subtotal
        fsliInfo.IsTotal = (InStr(1, fsliName, "total", vbTextCompare) > 0)
        fsliInfo.IsSubtotal = (InStr(1, fsliName, "subtotal", vbTextCompare) > 0) Or _
                              (InStr(1, fsliName, "sub-total", vbTextCompare) > 0)
        
        ' Detect level (indentation)
        fsliInfo.Level = DetectIndentationLevel(ws, row, 2)
        
        ' Add to collection
        fsliList.Add fsliInfo
        
NextRow:
    Next row
    
    Set AnalyzeFSLiStructure = fsliList
    Exit Function
    
ErrorHandler:
    MsgBox "Error analyzing FSLi structure: " & Err.Description, vbCritical
    Set AnalyzeFSLiStructure = New Collection
End Function

' Detect indentation level of a cell
Private Function DetectIndentationLevel(ws As Worksheet, row As Long, col As Long) As Long
    On Error Resume Next
    DetectIndentationLevel = ws.Cells(row, col).IndentLevel
    If Err.Number <> 0 Then DetectIndentationLevel = 0
    On Error GoTo 0
End Function

' Check if entire row is empty
Private Function IsRowEmpty(ws As Worksheet, row As Long) As Boolean
    Dim col As Long
    Dim lastCol As Long
    
    lastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    
    For col = 1 To lastCol
        If ws.Cells(row, col).Value <> "" Then
            IsRowEmpty = False
            Exit Function
        End If
    Next col
    
    IsRowEmpty = True
End Function

' Create Full Input Table
Private Sub CreateFullInputTable(sourceWs As Worksheet, columns As Collection, _
                                 fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler
    
    Dim outputWs As Worksheet
    Dim outRow As Long
    Dim outCol As Long
    Dim i As Long
    Dim j As Long
    Dim colInfo As ColumnInfo
    Dim fsliInfo As FSLiInfo
    Dim packList As Collection
    Dim packName As String
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "Full Input Table"
    
    ' Get list of packs with selected column type
    Set packList = New Collection
    For i = 1 To columns.Count
        colInfo = columns(i)
        If colInfo.ColumnType = columnType And colInfo.PackName <> "" Then
            ' Avoid duplicates
            On Error Resume Next
            packList.Add colInfo.PackName, colInfo.PackName
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' Write headers
    outputWs.Cells(1, 1).Value = "Pack"
    
    outCol = 2
    For i = 1 To fsliList.Count
        fsliInfo = fsliList(i)
        outputWs.Cells(1, outCol).Value = fsliInfo.FSLiName
        
        ' Add metadata
        If fsliInfo.IsTotal Then
            outputWs.Cells(1, outCol).Value = outputWs.Cells(1, outCol).Value & " (Total)"
        ElseIf fsliInfo.IsSubtotal Then
            outputWs.Cells(1, outCol).Value = outputWs.Cells(1, outCol).Value & " (Subtotal)"
        End If
        
        outCol = outCol + 1
    Next i
    
    ' Write pack names and data
    outRow = 2
    For i = 1 To packList.Count
        packName = packList(i)
        outputWs.Cells(outRow, 1).Value = packName
        
        ' For each FSLi, find the value
        outCol = 2
        For j = 1 To fsliList.Count
            fsliInfo = fsliList(j)
            
            ' Find column for this pack
            Dim packCol As Long
            packCol = FindPackColumn(columns, packName, columnType)
            
            If packCol > 0 Then
                ' Copy value
                outputWs.Cells(outRow, outCol).Value = sourceWs.Cells(fsliInfo.RowIndex, packCol).Value
            End If
            
            outCol = outCol + 1
        Next j
        
        outRow = outRow + 1
    Next i
    
    ' Format as table
    FormatAsTable outputWs
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Full Input Table: " & Err.Description, vbCritical
End Sub

' Find column index for a specific pack
Private Function FindPackColumn(columns As Collection, packName As String, columnType As String) As Long
    Dim i As Long
    Dim colInfo As ColumnInfo
    
    For i = 1 To columns.Count
        colInfo = columns(i)
        If colInfo.PackName = packName And colInfo.ColumnType = columnType Then
            FindPackColumn = colInfo.ColumnIndex
            Exit Function
        End If
    Next i
    
    FindPackColumn = 0
End Function

' Format worksheet as table
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
    End If
    
    On Error GoTo 0
End Sub

' Process Discontinued tab (similar structure to Input tab)
Private Sub ProcessDiscontinuedTab(ws As Worksheet)
    ' Implementation similar to ProcessInputTab
    ' Omitted for brevity - would follow same pattern
End Sub

' Process Journals tab (similar structure to Input tab)
Private Sub ProcessJournalsTab(ws As Worksheet)
    ' Implementation similar to ProcessInputTab
    ' Omitted for brevity - would follow same pattern
End Sub

' Process Console tab (similar structure to Input tab)
Private Sub ProcessConsoleTab(ws As Worksheet)
    ' Implementation similar to ProcessInputTab
    ' Omitted for brevity - would follow same pattern
End Sub

' Create FSLi Key Table (placeholder)
Private Sub CreateFSLiKeyTable()
    ' Implementation for FSLi Key Table
    ' Will be added in next module
End Sub

' Create Pack Number Company Table (placeholder)
Private Sub CreatePackNumberCompanyTable()
    ' Implementation for Pack Number Company Table
    ' Will be added in next module
End Sub

' Create Percentage Tables (placeholder)
Private Sub CreatePercentageTables()
    ' Implementation for Percentage Tables
    ' Will be added in next module
End Sub
