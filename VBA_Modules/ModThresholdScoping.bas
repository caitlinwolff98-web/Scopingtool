Attribute VB_Name = "ModThresholdScoping"
Option Explicit

' ============================================================================
' MODULE: ModThresholdScoping
' PURPOSE: Handle threshold-based automatic scoping
' DESCRIPTION: Prompts user for FSLI selection and thresholds, then
'              automatically marks packs as "Scoped In" based on criteria
' ============================================================================

' Structure to hold threshold configuration
Private Type ThresholdConfig
    FSLiName As String
    ThresholdValue As Double
    ThresholdType As String ' "Absolute" or "Percentage"
End Type

' Main function to configure and apply thresholds
Public Function ConfigureAndApplyThresholds() As Collection
    On Error GoTo ErrorHandler
    
    Dim fsliList As Collection
    Dim selectedFSLis As Collection
    Dim thresholds As Collection
    Dim i As Long
    
    ' Get list of all FSLIs from the input tab
    Set fsliList = GetAvailableFSLIs()
    
    If fsliList.Count = 0 Then
        MsgBox "No FSLIs found for threshold configuration." & vbCrLf & vbCrLf & _
               "This may indicate:" & vbCrLf & _
               "1. The Input Continuing Operations tab was not found" & vbCrLf & _
               "2. All FSLIs were filtered out as statement headers" & vbCrLf & _
               "3. The data starts at a different row than expected (Row 9)" & vbCrLf & vbCrLf & _
               "Please verify your consolidation workbook structure.", _
               vbExclamation, "No FSLIs Available"
        Set ConfigureAndApplyThresholds = New Collection
        Exit Function
    End If
    
    ' Prompt user to select FSLIs for threshold application
    Set selectedFSLis = PromptUserForFSLISelection(fsliList)
    
    If selectedFSLis.Count = 0 Then
        MsgBox "No FSLIs selected for threshold-based scoping.", vbInformation
        Set ConfigureAndApplyThresholds = New Collection
        Exit Function
    End If
    
    ' For each selected FSLI, prompt for threshold value
    Set thresholds = New Collection
    For i = 1 To selectedFSLis.Count
        Dim fsliName As String
        Dim threshold As Object
        
        fsliName = selectedFSLis(i)
        Set threshold = PromptUserForThreshold(fsliName)
        
        If Not threshold Is Nothing Then
            thresholds.Add threshold
        End If
    Next i
    
    ' Return the threshold configuration
    Set ConfigureAndApplyThresholds = thresholds
    Exit Function
    
ErrorHandler:
    MsgBox "Error in threshold configuration: " & Err.Description, vbCritical
    Set ConfigureAndApplyThresholds = New Collection
End Function

' Get list of available FSLIs from the source workbook
Private Function GetAvailableFSLIs() As Collection
    On Error GoTo ErrorHandler
    
    Dim fsliList As New Collection
    Dim inputTab As Worksheet
    Dim row As Long
    Dim lastRow As Long
    Dim fsliName As String
    Dim fsliDict As Object
    
    Set fsliDict = CreateObject("Scripting.Dictionary")
    
    ' Get the input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    
    If inputTab Is Nothing Then
        Set GetAvailableFSLIs = fsliList
        Exit Function
    End If
    
    ' Find last row
    lastRow = inputTab.Cells(inputTab.Rows.Count, 2).End(xlUp).Row
    
    ' Collect unique FSLIs (excluding headers and totals for threshold selection)
    For row = 9 To lastRow
        fsliName = Trim(inputTab.Cells(row, 2).Value)
        
        ' Skip empty, notes, and statement headers
        If fsliName <> "" And _
           UCase(fsliName) <> "NOTES" And _
           Not IsStatementHeader(fsliName) Then
            
            ' Skip if already in dictionary
            If Not fsliDict.Exists(fsliName) Then
                ' Add to dictionary and collection
                fsliDict.Add fsliName, True
                fsliList.Add fsliName
            End If
        End If
    Next row
    
    Set GetAvailableFSLIs = fsliList
    Exit Function
    
ErrorHandler:
    MsgBox "Error getting FSLIs: " & Err.Description, vbCritical
    Set GetAvailableFSLIs = New Collection
End Function

' Check if a line is a statement header (copied from ModDataProcessing)
Private Function IsStatementHeader(fsliName As String) As Boolean
    Dim upperName As String
    upperName = UCase(Trim(fsliName))
    
    IsStatementHeader = False
    
    If upperName = "INCOME STATEMENT" Or _
       upperName = "BALANCE SHEET" Or _
       upperName = "STATEMENT OF FINANCIAL POSITION" Or _
       upperName = "STATEMENT OF PROFIT OR LOSS" Or _
       upperName = "STATEMENT OF COMPREHENSIVE INCOME" Or _
       upperName = "CASH FLOW STATEMENT" Or _
       upperName = "STATEMENT OF CASH FLOWS" Or _
       upperName = "STATEMENT OF CHANGES IN EQUITY" Then
        IsStatementHeader = True
    End If
End Function

' Prompt user to select FSLIs using a custom dialog
Private Function PromptUserForFSLISelection(fsliList As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Dim selectedFSLis As New Collection
    Dim msg As String
    Dim i As Long
    Dim userInput As String
    Dim selectedIndices() As String
    Dim index As Variant
    Dim fsliName As String
    
    ' Build message with FSLI list
    msg = "THRESHOLD-BASED SCOPING CONFIGURATION" & vbCrLf & vbCrLf
    msg = msg & "Select FSLIs for automatic threshold-based scoping." & vbCrLf
    msg = msg & "Enter the numbers of FSLIs you want to apply thresholds to," & vbCrLf
    msg = msg & "separated by commas (e.g., 1,3,5)" & vbCrLf & vbCrLf
    msg = msg & "NOTE: You can select Balance Sheet items like 'Total Assets'" & vbCrLf
    msg = msg & "or Income Statement items like 'Revenue'. The list below" & vbCrLf
    msg = msg & "excludes only statement headers (e.g., 'BALANCE SHEET')." & vbCrLf & vbCrLf
    msg = msg & "Available FSLIs:" & vbCrLf
    msg = msg & String(50, "-") & vbCrLf
    
    ' List all FSLIs
    For i = 1 To fsliList.Count
        msg = msg & i & ". " & fsliList(i) & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Total FSLIs available: " & fsliList.Count & vbCrLf
    msg = msg & vbCrLf & "Enter selection:" & vbCrLf
    msg = msg & "• Numbers (e.g., 1,3,5) OR" & vbCrLf
    msg = msg & "• FSLi names (e.g., Total Assets, Revenue)" & vbCrLf
    msg = msg & "• Leave blank to skip" & vbCrLf
    
    ' Get user input
    userInput = InputBox(msg, "Select FSLIs for Threshold Scoping", "")
    
    ' Parse user input
    If Trim(userInput) = "" Then
        Set PromptUserForFSLISelection = selectedFSLis
        Exit Function
    End If
    
    ' Split by comma
    selectedIndices = Split(userInput, ",")
    
    ' Validate and add selected FSLIs
    For Each index In selectedIndices
        Dim idx As Long
        Dim trimmedInput As String
        trimmedInput = Trim(CStr(index))
        
        ' Check if it's a number or a name
        If IsNumeric(trimmedInput) Then
            ' It's a number - use index
            idx = Val(trimmedInput)
            
            If idx >= 1 And idx <= fsliList.Count Then
                fsliName = fsliList(idx)
                selectedFSLis.Add fsliName
            End If
        Else
            ' It's a name - search for it in the list
            Dim foundMatch As Boolean
            foundMatch = False
            
            For i = 1 To fsliList.Count
                If UCase(Trim(fsliList(i))) = UCase(trimmedInput) Then
                    selectedFSLis.Add fsliList(i)
                    foundMatch = True
                    Exit For
                End If
            Next i
            
            ' If no exact match, try partial match
            If Not foundMatch Then
                For i = 1 To fsliList.Count
                    If InStr(1, UCase(fsliList(i)), UCase(trimmedInput), vbTextCompare) > 0 Then
                        ' Show confirmation dialog for partial matches
                        Dim confirmResult As VbMsgBoxResult
                        confirmResult = MsgBox("Did you mean: " & fsliList(i) & "?", _
                                              vbYesNo + vbQuestion, "Confirm FSLi Selection")
                        
                        If confirmResult = vbYes Then
                            selectedFSLis.Add fsliList(i)
                            Exit For
                        End If
                    End If
                Next i
            End If
        End If
    Next index
    
    Set PromptUserForFSLISelection = selectedFSLis
    Exit Function
    
ErrorHandler:
    MsgBox "Error in FSLI selection: " & Err.Description, vbCritical
    Set PromptUserForFSLISelection = New Collection
End Function

' Prompt user for threshold value for a specific FSLI
Private Function PromptUserForThreshold(fsliName As String) As Object
    On Error GoTo ErrorHandler
    
    Dim threshold As Object
    Dim userInput As String
    Dim thresholdValue As Double
    Dim msg As String
    
    ' Build message
    msg = "THRESHOLD CONFIGURATION" & vbCrLf & vbCrLf
    msg = msg & "FSLI: " & fsliName & vbCrLf & vbCrLf
    msg = msg & "Enter the threshold value (absolute amount)." & vbCrLf
    msg = msg & "Packs exceeding this value will be automatically scoped in." & vbCrLf & vbCrLf
    msg = msg & "Example: 300000000 (for 300 million)" & vbCrLf
    msg = msg & "Example: 50000 (for 50 thousand)" & vbCrLf & vbCrLf
    msg = msg & "Enter threshold value:"
    
    ' Get user input
    userInput = InputBox(msg, "Enter Threshold for " & fsliName, "0")
    
    ' Validate input
    If Not IsNumeric(userInput) Then
        MsgBox "Invalid threshold value. Skipping this FSLI.", vbExclamation
        Set PromptUserForThreshold = Nothing
        Exit Function
    End If
    
    thresholdValue = CDbl(userInput)
    
    ' Create threshold object
    Set threshold = CreateObject("Scripting.Dictionary")
    threshold("FSLiName") = fsliName
    threshold("ThresholdValue") = thresholdValue
    threshold("ThresholdType") = "Absolute"
    
    Set PromptUserForThreshold = threshold
    Exit Function
    
ErrorHandler:
    MsgBox "Error configuring threshold: " & Err.Description, vbCritical
    Set PromptUserForThreshold = Nothing
End Function

' Apply thresholds to determine which packs should be scoped in
Public Function ApplyThresholdsToData(thresholds As Collection) As Object
    On Error GoTo ErrorHandler
    
    Dim scopedPacks As Object ' Dictionary: PackCode -> True
    Dim inputTab As Worksheet
    Dim threshold As Object
    Dim i As Long
    Dim row As Long
    Dim col As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim fsliName As String
    Dim packCode As String
    Dim cellValue As Variant
    
    Set scopedPacks = CreateObject("Scripting.Dictionary")
    
    ' Get input tab
    Set inputTab = ModTableGeneration.GetTabByCategory(ModConfig.CAT_INPUT_CONTINUING)
    If inputTab Is Nothing Then
        Set ApplyThresholdsToData = scopedPacks
        Exit Function
    End If
    
    ' Get dimensions
    lastRow = inputTab.Cells(inputTab.Rows.Count, 2).End(xlUp).Row
    lastCol = inputTab.Cells(7, inputTab.Columns.Count).End(xlToLeft).Column
    
    ' For each threshold configuration
    For i = 1 To thresholds.Count
        Set threshold = thresholds(i)
        fsliName = threshold("FSLiName")
        
        ' Find the row for this FSLI
        For row = 9 To lastRow
            If Trim(inputTab.Cells(row, 2).Value) = fsliName Then
                ' Check each pack (column) for this FSLI
                For col = 3 To lastCol ' Start from column 3 (after FSLI name columns)
                    cellValue = inputTab.Cells(row, col).Value
                    
                    If IsNumeric(cellValue) Then
                        ' Check if value exceeds threshold
                        If Abs(CDbl(cellValue)) >= threshold("ThresholdValue") Then
                            ' Get pack code from row 8
                            packCode = Trim(inputTab.Cells(8, col).Value)
                            
                            ' Exclude consolidated pack
                            If packCode <> "" And packCode <> g_ConsolidatedPackCode Then
                                ' Mark this pack as scoped in
                                If Not scopedPacks.Exists(packCode) Then
                                    scopedPacks.Add packCode, fsliName ' Store which FSLI triggered it
                                End If
                            End If
                        End If
                    End If
                Next col
                
                Exit For ' Found the FSLI row, move to next threshold
            End If
        Next row
    Next i
    
    Set ApplyThresholdsToData = scopedPacks
    Exit Function
    
ErrorHandler:
    MsgBox "Error applying thresholds: " & Err.Description, vbCritical
    Set ApplyThresholdsToData = CreateObject("Scripting.Dictionary")
End Function

' Create a sheet documenting the threshold configuration
Public Sub CreateThresholdConfigSheet(thresholds As Collection, scopedPacks As Object)
    On Error GoTo ErrorHandler
    
    Dim configWs As Worksheet
    Dim row As Long
    Dim threshold As Object
    Dim i As Long
    Dim packCode As Variant
    
    ' Check if sheet already exists
    On Error Resume Next
    Set configWs = g_OutputWorkbook.Worksheets("Threshold Configuration")
    On Error GoTo ErrorHandler
    
    If configWs Is Nothing Then
        Set configWs = g_OutputWorkbook.Worksheets.Add
        configWs.Name = "Threshold Configuration"
    Else
        configWs.Cells.Clear
    End If
    
    ' Write header
    row = 1
    With configWs
        .Cells(row, 1).Value = "THRESHOLD-BASED SCOPING CONFIGURATION"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 1).Font.Size = 14
        row = row + 2
        
        ' Thresholds section
        .Cells(row, 1).Value = "Configured Thresholds:"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "FSLI"
        .Cells(row, 2).Value = "Threshold Value"
        .Cells(row, 3).Value = "Type"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Font.Bold = True
        .Cells(row, 3).Font.Bold = True
        row = row + 1
        
        ' Write threshold details
        For i = 1 To thresholds.Count
            Set threshold = thresholds(i)
            .Cells(row, 1).Value = threshold("FSLiName")
            .Cells(row, 2).Value = threshold("ThresholdValue")
            .Cells(row, 2).NumberFormat = "#,##0"
            .Cells(row, 3).Value = threshold("ThresholdType")
            row = row + 1
        Next i
        
        ' Scoped packs section
        row = row + 2
        .Cells(row, 1).Value = "Packs Automatically Scoped In:"
        .Cells(row, 1).Font.Bold = True
        row = row + 1
        
        .Cells(row, 1).Value = "Pack Code"
        .Cells(row, 2).Value = "Triggered By FSLI"
        .Cells(row, 1).Font.Bold = True
        .Cells(row, 2).Font.Bold = True
        row = row + 1
        
        ' Write scoped packs
        For Each packCode In scopedPacks.Keys
            .Cells(row, 1).Value = packCode
            .Cells(row, 2).Value = scopedPacks(packCode)
            row = row + 1
        Next packCode
        
        ' Summary
        row = row + 2
        .Cells(row, 1).Value = "Total Packs Scoped In: " & scopedPacks.Count
        .Cells(row, 1).Font.Bold = True
        
        ' Auto-fit columns
        .Columns.AutoFit
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating threshold config sheet: " & Err.Description, vbCritical
End Sub
