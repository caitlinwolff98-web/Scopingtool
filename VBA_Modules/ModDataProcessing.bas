Option Explicit

' ============================================================================
' MODULE: ModDataProcessing
' PURPOSE: Process consolidation data and analyze structure
' DESCRIPTION: Handles unmerging cells, detecting columns, analyzing FSLi
'              hierarchies, and extracting entity information
' ============================================================================

' Main processing orchestrator
Public Sub ProcessConsolidationData()
    On Error GoTo ErrorHandler
    
    Dim inputTab As Worksheet
    Dim discontinuedTab As Worksheet
    Dim journalsTab As Worksheet
    Dim consoleTab As Worksheet
    
    ' Get required tabs
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    
    If inputTab Is Nothing Then
        MsgBox "Could not find Input Continuing tab. Process cannot continue.", vbCritical
        Exit Sub
    End If
    
    ' Process Input Continuing tab
    Application.StatusBar = "Processing Input Continuing tab..."
    ProcessInputTab inputTab
    
    ' Process other tabs if they exist
    Set discontinuedTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_DISCONTINUED)
    If Not discontinuedTab Is Nothing Then
        Application.StatusBar = "Processing Discontinued tab..."
        ProcessDiscontinuedTab discontinuedTab
    End If
    
    Set journalsTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_JOURNALS_CONTINUING)
    If Not journalsTab Is Nothing Then
        Application.StatusBar = "Processing Journals tab..."
        ProcessJournalsTab journalsTab
    End If
    
    Set consoleTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_CONSOLE_CONTINUING)
    If Not consoleTab Is Nothing Then
        Application.StatusBar = "Processing Consol tab..."
        ProcessConsolTab consoleTab
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
    Dim colInfo As Object ' Use dictionary instead of UDT
    Dim packName As String
    Dim packCode As String
    
    ' Find last column with data in row 6
    lastCol = ws.Cells(6, ws.columns.count).End(xlToLeft).Column
    
    ' Analyze row 6 for column types
    For col = 1 To lastCol
        cellValue = Trim(ws.Cells(6, col).Value)
        
        If cellValue <> "" Then
            Set colInfo = CreateObject("Scripting.Dictionary")
            colInfo("ColumnIndex") = col
            
            ' Determine column type
            If InStr(1, cellValue, "original", vbTextCompare) > 0 And _
               InStr(1, cellValue, "entity currency", vbTextCompare) > 0 Then
                colInfo("ColumnType") = "Original/Entity"
            ElseIf InStr(1, cellValue, "consolidation", vbTextCompare) > 0 And _
                   InStr(1, cellValue, "consolidation currency", vbTextCompare) > 0 Then
                colInfo("ColumnType") = "Consolidation/Consolidation"
            Else
                colInfo("ColumnType") = "Other"
            End If
            
            ' Get pack name from row 7
            packName = ""
            If ws.Cells(7, col).Value <> "" Then
                packName = Trim(ws.Cells(7, col).Value)
            End If
            colInfo("PackName") = packName
            
            ' Get pack code from row 8
            packCode = ""
            If ws.Cells(8, col).Value <> "" Then
                packCode = Trim(ws.Cells(8, col).Value)
            End If
            colInfo("PackCode") = packCode
            
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
    Dim colInfo As Object
    Dim msg As String
    Dim response As VbMsgBoxResult
    
    ' Count column types
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = "Original/Entity" Then
            originalCount = originalCount + 1
        ElseIf colInfo("ColumnType") = "Consolidation/Consolidation" Then
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
    Dim fsliInfo As Object ' Dictionary instead of UDT
    Dim currentStatement As String
    Dim notesStartRow As Long
    
    currentStatement = ""
    notesStartRow = 0
    
    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    
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
        
        ' Detect statement type and skip statement headers
        If InStr(1, fsliName, "income statement", vbTextCompare) > 0 Then
            currentStatement = "Income Statement"
            ' Skip if this is ONLY a statement header (not a line item)
            If IsStatementHeader(fsliName) Then
                GoTo NextRow
            End If
        ElseIf InStr(1, fsliName, "balance sheet", vbTextCompare) > 0 Then
            currentStatement = "Balance Sheet"
            ' Skip if this is ONLY a statement header (not a line item)
            If IsStatementHeader(fsliName) Then
                GoTo NextRow
            End If
        End If
        
        ' Skip other common header patterns
        If IsStatementHeader(fsliName) Then
            GoTo NextRow
        End If
        
        ' Create FSLi info dictionary
        Set fsliInfo = CreateObject("Scripting.Dictionary")
        fsliInfo("FSLiName") = fsliName
        fsliInfo("RowIndex") = row
        fsliInfo("StatementType") = currentStatement
        
        ' Detect if it's a total or subtotal
        fsliInfo("IsTotal") = (InStr(1, fsliName, "total", vbTextCompare) > 0)
        fsliInfo("IsSubtotal") = (InStr(1, fsliName, "subtotal", vbTextCompare) > 0) Or _
                                 (InStr(1, fsliName, "sub-total", vbTextCompare) > 0)
        
        ' Detect level (indentation)
        fsliInfo("Level") = DetectIndentationLevel(ws, row, 2)
        
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
    If Err.Number <> 0 Then
        DetectIndentationLevel = 0
    End If
    On Error GoTo 0
End Function

' Check if a line is a statement header (not an actual FSLI)
Private Function IsStatementHeader(fsliName As String) As Boolean
    Dim upperName As String
    upperName = UCase(Trim(fsliName))
    
    ' Common statement headers to exclude
    IsStatementHeader = False
    
    ' Exact matches for statement headers
    If upperName = "INCOME STATEMENT" Or _
       upperName = "BALANCE SHEET" Or _
       upperName = "STATEMENT OF FINANCIAL POSITION" Or _
       upperName = "STATEMENT OF PROFIT OR LOSS" Or _
       upperName = "STATEMENT OF COMPREHENSIVE INCOME" Or _
       upperName = "CASH FLOW STATEMENT" Or _
       upperName = "STATEMENT OF CASH FLOWS" Or _
       upperName = "STATEMENT OF CHANGES IN EQUITY" Then
        IsStatementHeader = True
        Exit Function
    End If
    
    ' Check if it's a pure header without additional detail
    ' (e.g., "INCOME STATEMENT" yes, "INCOME STATEMENT - Revenue" no)
    If Len(upperName) < 50 Then ' Headers are typically short
        ' Check for statement indicators without line item details
        If (upperName = "INCOME STATEMENT" Or upperName = "BALANCE SHEET") And _
           InStr(upperName, "-") = 0 And _
           InStr(upperName, ":") = 0 And _
           InStr(upperName, "TOTAL") = 0 Then
            IsStatementHeader = True
            Exit Function
        End If
    End If
End Function

' Check if entire row is empty
Private Function IsRowEmpty(ws As Worksheet, row As Long) As Boolean
    Dim col As Long
    Dim lastCol As Long
    
    lastCol = ws.Cells(row, ws.columns.count).End(xlToLeft).Column
    
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
    
    CreateGenericTable sourceWs, columns, fsliList, columnType, "Full Input Table"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Full Input Table: " & Err.Description, vbCritical
End Sub

' Find column index for a specific pack
Private Function FindPackColumn(columns As Collection, packName As String, columnType As String) As Long
    Dim i As Long
    Dim colInfo As Object
    
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("PackName") = packName And colInfo("ColumnType") = columnType Then
            FindPackColumn = colInfo("ColumnIndex")
            Exit Function
        End If
    Next i
    
    FindPackColumn = 0
End Function

' Process Discontinued tab
Private Sub ProcessDiscontinuedTab(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String
    
    ' Unmerge all cells
    ws.Cells.UnMerge
    
    ' Detect columns
    Set columns = DetectColumns(ws)
    
    ' Use same column type as Input tab
    selectedColumnType = "Consolidation/Consolidation"
    
    ' Check if we have columns of this type
    Dim hasColumns As Boolean
    hasColumns = False
    Dim i As Long
    Dim colInfo As Object
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = selectedColumnType Then
            hasColumns = True
            Exit For
        End If
    Next i
    
    If Not hasColumns Then
        selectedColumnType = "Original/Entity"
    End If
    
    ' Analyze FSLi structure
    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    
    ' Create Discontinued Table
    CreateDiscontinuedTable ws, columns, fsliList, selectedColumnType
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing Discontinued tab: " & Err.Description, vbCritical
End Sub

' Process Journals tab
Private Sub ProcessJournalsTab(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String
    
    ' Unmerge all cells
    ws.Cells.UnMerge
    
    ' Detect columns
    Set columns = DetectColumns(ws)
    
    ' Use same column type as Input tab
    selectedColumnType = "Consolidation/Consolidation"
    
    ' Check if we have columns of this type
    Dim hasColumns As Boolean
    hasColumns = False
    Dim i As Long
    Dim colInfo As Object
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = selectedColumnType Then
            hasColumns = True
            Exit For
        End If
    Next i
    
    If Not hasColumns Then
        selectedColumnType = "Original/Entity"
    End If
    
    ' Analyze FSLi structure
    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    
    ' Create Journals Table
    CreateJournalsTable ws, columns, fsliList, selectedColumnType
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing Journals tab: " & Err.Description, vbCritical
End Sub

' Process Consol tab
Private Sub ProcessConsolTab(ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim columns As Collection
    Dim fsliList As Collection
    Dim selectedColumnType As String
    
    ' Unmerge all cells
    ws.Cells.UnMerge
    
    ' Detect columns
    Set columns = DetectColumns(ws)
    
    ' Use same column type as Input tab
    selectedColumnType = "Consolidation/Consolidation"
    
    ' Check if we have columns of this type
    Dim hasColumns As Boolean
    hasColumns = False
    Dim i As Long
    Dim colInfo As Object
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = selectedColumnType Then
            hasColumns = True
            Exit For
        End If
    Next i
    
    If Not hasColumns Then
        selectedColumnType = "Original/Entity"
    End If
    
    ' Analyze FSLi structure
    Set fsliList = AnalyzeFSLiStructure(ws, selectedColumnType)
    
    ' Create Consol Table
    CreateConsolTable ws, columns, fsliList, selectedColumnType
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error processing Consol tab: " & Err.Description, vbCritical
End Sub

' Create Journals Table
Private Sub CreateJournalsTable(sourceWs As Worksheet, columns As Collection, _
                                fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler
    
    CreateGenericTable sourceWs, columns, fsliList, columnType, "Journals Table"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Journals Table: " & Err.Description, vbCritical
End Sub

' Create Consol Table
Private Sub CreateConsolTable(sourceWs As Worksheet, columns As Collection, _
                               fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler
    
    CreateGenericTable sourceWs, columns, fsliList, columnType, "Full Consol Table"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Consol Table: " & Err.Description, vbCritical
End Sub

' Create Discontinued Table
Private Sub CreateDiscontinuedTable(sourceWs As Worksheet, columns As Collection, _
                                    fsliList As Collection, columnType As String)
    On Error GoTo ErrorHandler
    
    CreateGenericTable sourceWs, columns, fsliList, columnType, "Discontinued Table"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating Discontinued Table: " & Err.Description, vbCritical
End Sub

' Generic table creation function
Private Sub CreateGenericTable(sourceWs As Worksheet, columns As Collection, _
                               fsliList As Collection, columnType As String, tableName As String)
    On Error GoTo ErrorHandler
    
    Dim outputWs As Worksheet
    Dim outRow As Long
    Dim outCol As Long
    Dim i As Long
    Dim j As Long
    Dim colInfo As Object
    Dim fsliInfo As Object
    Dim packList As Collection
    Dim packName As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = tableName
    
    ' Get list of packs with selected column type
    Set packList = New Collection
    For i = 1 To columns.count
        Set colInfo = columns(i)
        If colInfo("ColumnType") = columnType And colInfo("PackName") <> "" Then
            ' Avoid duplicates
            On Error Resume Next
            packList.Add colInfo("PackName"), colInfo("PackName")
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' Write headers
    outputWs.Cells(1, 1).Value = "Pack"
    
    outCol = 2
    For i = 1 To fsliList.count
        Set fsliInfo = fsliList(i)
        outputWs.Cells(1, outCol).Value = fsliInfo("FSLiName")
        outCol = outCol + 1
    Next i
    
    ' Write pack names and data
    outRow = 2
    For i = 1 To packList.count
        packName = packList(i)
        outputWs.Cells(outRow, 1).Value = packName
        
        ' For each FSLi, find the value
        outCol = 2
        For j = 1 To fsliList.count
            Set fsliInfo = fsliList(j)
            
            ' Find column for this pack
            Dim packCol As Long
            packCol = FindPackColumn(columns, packName, columnType)
            
            If packCol > 0 Then
                ' Copy value
                outputWs.Cells(outRow, outCol).Value = sourceWs.Cells(fsliInfo("RowIndex"), packCol).Value
            End If
            
            outCol = outCol + 1
        Next j
        
        outRow = outRow + 1
    Next i
    
    ' Get dimensions for table
    lastRow = outputWs.Cells(outputWs.Rows.count, 1).End(xlUp).row
    lastCol = outputWs.Cells(1, outputWs.columns.count).End(xlToLeft).Column
    
    ' Create actual Excel Table
    If lastRow > 1 And lastCol > 1 Then
        On Error Resume Next
        Set tbl = outputWs.ListObjects.Add(xlSrcRange, outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = Replace(tableName, " ", "_")
            tbl.TableStyle = "TableStyleMedium2"
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Auto-fit columns
    outputWs.columns.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating " & tableName & ": " & Err.Description, vbCritical
End Sub

' Create FSLi Key Table
Private Sub CreateFSLiKeyTable()
    ' Call the implementation in ModTableGeneration
    ModTableGeneration.CreateFSLiKeyTable
End Sub

' Create Pack Number Company Table
Private Sub CreatePackNumberCompanyTable()
    ' Call the implementation in ModTableGeneration
    ModTableGeneration.CreatePackNumberCompanyTable
End Sub

' Create Percentage Tables
Private Sub CreatePercentageTables()
    ' Call the implementation in ModTableGeneration
    ModTableGeneration.CreatePercentageTables
End Sub

