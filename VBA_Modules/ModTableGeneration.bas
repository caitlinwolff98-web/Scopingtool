Option Explicit

' ============================================================================
' MODULE: ModTableGeneration
' PURPOSE: Generate supporting tables (FSLi Key, Pack Number, Percentages)
' DESCRIPTION: Creates additional tables required for Power BI integration
' ============================================================================

' Import category constants from ModTabCategorization
Public Const CAT_SEGMENT = "TGK Segment Tabs"
Public Const CAT_DISCONTINUED = "Discontinued Ops Tab"
Public Const CAT_INPUT_CONTINUING = "TGK Input Continuing Operations Tab"
Public Const CAT_JOURNALS_CONTINUING = "TGK Journals Continuing Tab"
Public Const CAT_CONSOLE_CONTINUING = "TGK Consol Continuing Tab"

' Get worksheet by category (moved here for accessibility)
Public Function GetTabByCategory(categoryName As String) As Worksheet
    On Error Resume Next
    Dim tabInfo As Object
    
    If g_TabCategories.Exists(categoryName) Then
        If g_TabCategories(categoryName).count > 0 Then
            Set tabInfo = g_TabCategories(categoryName)(1)
            Set GetTabByCategory = g_SourceWorkbook.Worksheets(tabInfo("TabName"))
        End If
    End If
    
    On Error GoTo 0
End Function

' Create FSLi Key Table with all FSLi entries
Public Sub CreateFSLiKeyTable()
    On Error GoTo ErrorHandler
    
    Dim outputWs As Worksheet
    Dim fsliCollection As Collection
    Dim fsliName As String
    Dim row As Long
    Dim i As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim tbl As ListObject
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = "FSLi Key Table"
    
    ' Set up headers
    outputWs.Cells(1, 1).Value = "FSLi"
    outputWs.Cells(1, 2).Value = "Statement Type"
    outputWs.Cells(1, 3).Value = "Is Total"
    outputWs.Cells(1, 4).Value = "Level"
    
    ' Get unique FSLi names from all tables
    Set fsliCollection = CollectAllFSLiNames()
    
    ' Populate FSLi names
    row = 2
    For i = 1 To fsliCollection.count
        Dim fsliDict As Object
        Set fsliDict = fsliCollection(i)
        
        outputWs.Cells(row, 1).Value = fsliDict("FSLiName")
        outputWs.Cells(row, 2).Value = fsliDict("StatementType")
        outputWs.Cells(row, 3).Value = IIf(fsliDict("IsTotal"), "Yes", "No")
        outputWs.Cells(row, 4).Value = fsliDict("Level")
        
        row = row + 1
    Next i
    
    ' Get dimensions
    lastRow = outputWs.Cells(outputWs.Rows.count, 1).End(xlUp).row
    lastCol = 4
    
    ' Create actual Excel Table
    If lastRow > 1 Then
        On Error Resume Next
        Set tbl = outputWs.ListObjects.Add(xlSrcRange, outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = "FSLi_Key_Table"
            tbl.TableStyle = "TableStyleMedium2"
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Auto-fit columns
    outputWs.columns.AutoFit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating FSLi Key Table: " & Err.Description, vbCritical
End Sub

' Collect all unique FSLi names from all tables with metadata
Private Function CollectAllFSLiNames() As Collection
    Dim fsliDict As Object ' Dictionary to track unique FSLi
    Dim ws As Worksheet
    Dim tab As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim fsliName As String
    Dim resultCollection As New Collection
    Dim fsliInfo As Object
    
    Set fsliDict = CreateObject("Scripting.Dictionary")
    
    ' Get FSLi info from source tabs
    Dim inputTab As Worksheet
    Set inputTab = GetTabByCategory(CAT_INPUT_CONTINUING)
    
    If Not inputTab Is Nothing Then
        ' Analyze FSLi structure from input tab
        lastRow = inputTab.Cells(inputTab.Rows.count, 2).End(xlUp).row
        
        For row = 9 To lastRow
            fsliName = Trim(inputTab.Cells(row, 2).Value)
            
            If fsliName <> "" And UCase(fsliName) <> "NOTES" Then
                If Not fsliDict.Exists(fsliName) Then
                    Set fsliInfo = CreateObject("Scripting.Dictionary")
                    fsliInfo("FSLiName") = fsliName
                    
                    ' Detect statement type
                    If InStr(1, fsliName, "income statement", vbTextCompare) > 0 Then
                        fsliInfo("StatementType") = "Income Statement"
                    ElseIf InStr(1, fsliName, "balance sheet", vbTextCompare) > 0 Then
                        fsliInfo("StatementType") = "Balance Sheet"
                    Else
                        fsliInfo("StatementType") = ""
                    End If
                    
                    ' Detect if it's a total
                    fsliInfo("IsTotal") = (InStr(1, fsliName, "total", vbTextCompare) > 0)
                    
                    ' Detect level
                    On Error Resume Next
                    fsliInfo("Level") = inputTab.Cells(row, 2).IndentLevel
                    If Err.Number <> 0 Then
                        fsliInfo("Level") = 0
                    End If
                    On Error GoTo 0
                    
                    fsliDict.Add fsliName, fsliInfo
                    resultCollection.Add fsliInfo
                End If
            End If
        Next row
    End If
    
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
    Dim lastRow As Long
    Dim tbl As ListObject
    
    Set packDict = CreateObject("Scripting.Dictionary")
    
    ' Get segment tabs
    Set segmentTabs = GetTabsForCategory(CAT_SEGMENT)
    
    ' Process each segment tab
    For i = 1 To segmentTabs.count
        Set tabInfo = segmentTabs(i)
        Set ws = g_SourceWorkbook.Worksheets(tabInfo("TabName"))
        
        ' Get division name
        divisionName = tabInfo("DivisionName")
        If divisionName = "" Then
            divisionName = PromptForDivisionName(tabInfo("TabName"))
            tabInfo("DivisionName") = divisionName
        End If
        
        ' Extract pack names from row 7 and codes from row 8
        lastCol = ws.Cells(7, ws.columns.count).End(xlToLeft).Column
        
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
    
    ' Also process Input, Journals, Console tabs
    Dim inputTab As Worksheet
    Dim journalsTab As Worksheet
    Dim consoleTab As Worksheet
    
    Set inputTab = GetTabByCategory(CAT_INPUT_CONTINUING)
    If Not inputTab Is Nothing Then
        lastCol = inputTab.Cells(7, inputTab.columns.count).End(xlToLeft).Column
        For col = 1 To lastCol
            packName = Trim(inputTab.Cells(7, col).Value)
            packCode = Trim(inputTab.Cells(8, col).Value)
            
            If packName <> "" And packCode <> "" Then
                packKey = packCode
                If Not packDict.Exists(packKey) Then
                    Set packInfo = CreateObject("Scripting.Dictionary")
                    packInfo("Name") = packName
                    packInfo("Code") = packCode
                    packInfo("Division") = "Continuing Operations"
                    packDict.Add packKey, packInfo
                End If
            End If
        Next col
    End If
    
    Set journalsTab = GetTabByCategory(CAT_JOURNALS_CONTINUING)
    If Not journalsTab Is Nothing Then
        lastCol = journalsTab.Cells(7, journalsTab.columns.count).End(xlToLeft).Column
        For col = 1 To lastCol
            packName = Trim(journalsTab.Cells(7, col).Value)
            packCode = Trim(journalsTab.Cells(8, col).Value)
            
            If packName <> "" And packCode <> "" Then
                packKey = packCode
                If Not packDict.Exists(packKey) Then
                    Set packInfo = CreateObject("Scripting.Dictionary")
                    packInfo("Name") = packName
                    packInfo("Code") = packCode
                    packInfo("Division") = "Journals"
                    packDict.Add packKey, packInfo
                End If
            End If
        Next col
    End If
    
    Set consoleTab = GetTabByCategory(CAT_CONSOLE_CONTINUING)
    If Not consoleTab Is Nothing Then
        lastCol = consoleTab.Cells(7, consoleTab.columns.count).End(xlToLeft).Column
        For col = 1 To lastCol
            packName = Trim(consoleTab.Cells(7, col).Value)
            packCode = Trim(consoleTab.Cells(8, col).Value)
            
            If packName <> "" And packCode <> "" Then
                packKey = packCode
                If Not packDict.Exists(packKey) Then
                    Set packInfo = CreateObject("Scripting.Dictionary")
                    packInfo("Name") = packName
                    packInfo("Code") = packCode
                    packInfo("Division") = "Consolidated"
                    packDict.Add packKey, packInfo
                End If
            End If
        Next col
    End If
    
    ' Process discontinued tab if it exists
    Set discontinuedTab = GetTabByCategory(CAT_DISCONTINUED)
    If Not discontinuedTab Is Nothing Then
        lastCol = discontinuedTab.Cells(7, discontinuedTab.columns.count).End(xlToLeft).Column
        
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
    
    ' Get dimensions
    lastRow = outputWs.Cells(outputWs.Rows.count, 1).End(xlUp).row
    
    ' Create actual Excel Table
    If lastRow > 1 Then
        On Error Resume Next
        Set tbl = outputWs.ListObjects.Add(xlSrcRange, outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, 3)), , xlYes)
        If Not tbl Is Nothing Then
            tbl.Name = "Pack_Number_Company_Table"
            tbl.TableStyle = "TableStyleMedium2"
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Auto-fit columns
    outputWs.columns.AutoFit
    
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
    Dim consolPackRow As Long
    Dim percentValue As Double
    Dim consolValue As Double
    Dim tbl As ListObject
    
    ' Create percentage table name
    percentTableName = Replace(sourceWs.Name, "Table", "Percentage")
    
    ' Create output worksheet
    Set outputWs = g_OutputWorkbook.Worksheets.Add
    outputWs.Name = percentTableName
    
    ' Get dimensions
    lastRow = sourceWs.Cells(sourceWs.Rows.count, 1).End(xlUp).row
    lastCol = sourceWs.Cells(1, sourceWs.columns.count).End(xlToLeft).Column
    
    ' Copy headers
    sourceWs.Rows(1).Copy outputWs.Rows(1)
    
    ' Find "The Bidvest Group Consolidated" row
    consolPackRow = 0
    For row = 2 To lastRow
        If InStr(1, sourceWs.Cells(row, 1).Value, "Bidvest Group Consolidated", vbTextCompare) > 0 Or _
           InStr(1, sourceWs.Cells(row, 1).Value, "The Bidvest Group Consolidated", vbTextCompare) > 0 Then
            consolPackRow = row
            Exit For
        End If
    Next row
    
    ' If consolidated pack not found, use column totals approach
    If consolPackRow = 0 Then
        ' Calculate percentages based on column totals
        For col = 2 To lastCol
            ' Calculate column total
            Dim columnTotal As Double
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
                    outputWs.Cells(row, col).Value = percentValue / 100 ' Excel percentage format
                Else
                    outputWs.Cells(row, col).Value = 0
                End If
                outputWs.Cells(row, col).NumberFormat = "0.00%"
            Next row
        Next col
    Else
        ' Calculate percentages based on consolidated pack
        For col = 2 To lastCol
            ' Get consolidated pack value for this FSLi
            consolValue = 0
            If IsNumeric(sourceWs.Cells(consolPackRow, col).Value) Then
                consolValue = Abs(sourceWs.Cells(consolPackRow, col).Value)
            End If
            
            ' Calculate percentage for each pack
            For row = 2 To lastRow
                cellValue = sourceWs.Cells(row, col).Value
                
                If IsNumeric(cellValue) And consolValue <> 0 Then
                    percentValue = (Abs(cellValue) / consolValue) * 100
                    outputWs.Cells(row, col).Value = percentValue / 100 ' Excel percentage format
                ElseIf IsNumeric(cellValue) And consolValue = 0 Then
                    outputWs.Cells(row, col).Value = 0
                Else
                    outputWs.Cells(row, col).Value = 0
                End If
                outputWs.Cells(row, col).NumberFormat = "0.00%"
            Next row
        Next col
    End If
    
    ' Copy pack names
    For row = 2 To lastRow
        outputWs.Cells(row, 1).Value = sourceWs.Cells(row, 1).Value
    Next row
    
    ' Create actual Excel Table
    On Error Resume Next
    Set tbl = outputWs.ListObjects.Add(xlSrcRange, outputWs.Range(outputWs.Cells(1, 1), outputWs.Cells(lastRow, lastCol)), , xlYes)
    If Not tbl Is Nothing Then
        tbl.Name = Replace(percentTableName, " ", "_")
        tbl.TableStyle = "TableStyleMedium2"
    End If
    On Error GoTo ErrorHandler
    
    ' Auto-fit columns
    outputWs.columns.AutoFit
    
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
    Dim tbl As ListObject
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.columns.count).End(xlToLeft).Column
    
    If lastRow > 1 And lastCol > 1 Then
        ' Create Excel Table object if not already created
        If ws.ListObjects.count = 0 Then
            Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
            If Not tbl Is Nothing Then
                tbl.Name = Replace(ws.Name, " ", "_")
                tbl.TableStyle = "TableStyleMedium2"
            End If
        End If
        
        ' Auto-fit columns
        ws.columns.AutoFit
        
        ' Freeze top row
        ws.Activate
        ws.Rows(2).Select
        ActiveWindow.FreezePanes = True
        ws.Range("A1").Select
    End If
    
    On Error GoTo 0
End Sub

