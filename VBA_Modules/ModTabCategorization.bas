Attribute VB_Name = "ModTabCategorization"
Option Explicit

' ============================================================================
' MODULE: ModTabCategorization
' PURPOSE: Handle tab categorization and validation
' DESCRIPTION: Provides functionality to categorize worksheets into predefined
'              categories for proper processing
' ============================================================================

' Category constants
Public Const CAT_SEGMENT = "TGK Segment Tabs"
Public Const CAT_DISCONTINUED = "TGK Discontinued Opt Tab"
Public Const CAT_INPUT_CONTINUING = "TGK Input Continuing Operations Tab"
Public Const CAT_JOURNALS_CONTINUING = "TGK Journals Continuing Tab"
Public Const CAT_CONSOLE_CONTINUING = "TGK Console Continuing Tab"
Public Const CAT_BS = "TGK BS Tab"
Public Const CAT_IS = "TGK IS Tab"
Public Const CAT_PULL_WORKINGS = "Pull Workings"
Public Const CAT_UNCATEGORIZED = "Uncategorized"

' Structure to hold tab categorization
Private Type TabCategory
    TabName As String
    Category As String
    DivisionName As String ' Used for segment tabs
End Type

Private m_TabCategories() As TabCategory
Private m_TabCount As Long

' Initialize the categorization system
Public Function CategorizeTabs(tabList As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim userForm As Object
    
    ' Initialize array
    m_TabCount = tabList.Count
    ReDim m_TabCategories(1 To m_TabCount)
    
    ' Populate tab names
    For i = 1 To tabList.Count
        m_TabCategories(i).TabName = tabList(i)
        m_TabCategories(i).Category = CAT_UNCATEGORIZED
        m_TabCategories(i).DivisionName = ""
    Next i
    
    ' Show categorization interface
    If Not ShowCategorizationDialog() Then
        CategorizeTabs = False
        Exit Function
    End If
    
    ' Store categorization in global dictionary
    Set g_TabCategories = CreateObject("Scripting.Dictionary")
    
    For i = 1 To m_TabCount
        If Not g_TabCategories.Exists(m_TabCategories(i).Category) Then
            g_TabCategories.Add m_TabCategories(i).Category, New Collection
        End If
        
        Dim tabInfo As Object
        Set tabInfo = CreateObject("Scripting.Dictionary")
        tabInfo("TabName") = m_TabCategories(i).TabName
        tabInfo("DivisionName") = m_TabCategories(i).DivisionName
        
        g_TabCategories(m_TabCategories(i).Category).Add tabInfo
    Next i
    
    CategorizeTabs = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in tab categorization: " & Err.Description, vbCritical
    CategorizeTabs = False
End Function

' Show dialog for categorizing tabs
Private Function ShowCategorizationDialog() As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim response As VbMsgBoxResult
    Dim categoryChoice As String
    Dim validCategories As String
    
    ' Create temporary worksheet for categorization
    Set ws = g_SourceWorkbook.Worksheets.Add
    ws.Name = "TempCategorization_" & Format(Now, "hhmmss")
    
    ' Set up headers
    ws.Range("A1").Value = "Tab Name"
    ws.Range("B1").Value = "Category"
    ws.Range("C1").Value = "Division Name (for segments)"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(200, 200, 200)
    
    ' List all tabs
    For i = 1 To m_TabCount
        ws.Cells(i + 1, 1).Value = m_TabCategories(i).TabName
        ws.Cells(i + 1, 2).Value = CAT_UNCATEGORIZED
    Next i
    
    ' Add validation list for categories
    validCategories = CAT_SEGMENT & "," & CAT_DISCONTINUED & "," & CAT_INPUT_CONTINUING & "," & _
                      CAT_JOURNALS_CONTINUING & "," & CAT_CONSOLE_CONTINUING & "," & _
                      CAT_BS & "," & CAT_IS & "," & CAT_PULL_WORKINGS & "," & CAT_UNCATEGORIZED
    
    With ws.Range("B2:B" & (m_TabCount + 1)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=validCategories
    End With
    
    ' Auto-fit columns
    ws.Columns("A:C").AutoFit
    
    ' Instructions
    MsgBox "Please categorize the tabs in the newly created worksheet." & vbCrLf & vbCrLf & _
           "Instructions:" & vbCrLf & _
           "1. Select a category from the dropdown in column B for each tab" & vbCrLf & _
           "2. For segment tabs, enter the division name in column C" & vbCrLf & _
           "3. Categories marked with (*) can only have ONE tab" & vbCrLf & _
           "4. When finished, click OK on the next dialog" & vbCrLf & vbCrLf & _
           "Category Descriptions:" & vbCrLf & _
           "- " & CAT_SEGMENT & ": All segment tabs (multiple allowed)" & vbCrLf & _
           "- " & CAT_DISCONTINUED & ": Discontinued operations (*)" & vbCrLf & _
           "- " & CAT_INPUT_CONTINUING & ": Input continuing operations (*)" & vbCrLf & _
           "- " & CAT_JOURNALS_CONTINUING & ": Journal entries (*)" & vbCrLf & _
           "- " & CAT_CONSOLE_CONTINUING & ": Consolidated continuing (*)" & vbCrLf & _
           "- " & CAT_BS & ": Balance sheet (*)" & vbCrLf & _
           "- " & CAT_IS & ": Income statement (*)" & vbCrLf & _
           "- " & CAT_PULL_WORKINGS & ": Working papers (multiple allowed)", _
           vbInformation, "Categorize Tabs"
    
    ' Wait for user to finish
    ws.Activate
    response = MsgBox("Have you finished categorizing all tabs?", vbYesNo + vbQuestion, "Confirm Categorization")
    
    If response = vbNo Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        ShowCategorizationDialog = False
        Exit Function
    End If
    
    ' Read categorization
    For i = 1 To m_TabCount
        m_TabCategories(i).Category = ws.Cells(i + 1, 2).Value
        m_TabCategories(i).DivisionName = ws.Cells(i + 1, 3).Value
    Next i
    
    ' Validate single-tab categories
    If Not ValidateSingleTabCategories() Then
        ShowCategorizationDialog = False
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        Exit Function
    End If
    
    ' Show uncategorized tabs
    ShowUncategorizedTabs
    
    ' Clean up
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    ShowCategorizationDialog = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in categorization dialog: " & Err.Description, vbCritical
    ShowCategorizationDialog = False
End Function

' Validate that single-tab categories have only one tab
Private Function ValidateSingleTabCategories() As Boolean
    Dim singleCategories As Variant
    Dim cat As Variant
    Dim count As Long
    Dim i As Long
    Dim msg As String
    
    singleCategories = Array(CAT_DISCONTINUED, CAT_INPUT_CONTINUING, CAT_JOURNALS_CONTINUING, _
                            CAT_CONSOLE_CONTINUING, CAT_BS, CAT_IS)
    
    For Each cat In singleCategories
        count = 0
        For i = 1 To m_TabCount
            If m_TabCategories(i).Category = cat Then
                count = count + 1
            End If
        Next i
        
        If count > 1 Then
            msg = "Category '" & cat & "' can only have ONE tab assigned, but " & count & " tabs were assigned." & vbCrLf & _
                  "Please correct this and try again."
            MsgBox msg, vbExclamation, "Validation Error"
            ValidateSingleTabCategories = False
            Exit Function
        End If
    Next cat
    
    ValidateSingleTabCategories = True
End Function

' Show uncategorized tabs to user
Private Sub ShowUncategorizedTabs()
    Dim i As Long
    Dim uncategorizedList As String
    Dim count As Long
    Dim response As VbMsgBoxResult
    
    uncategorizedList = ""
    count = 0
    
    For i = 1 To m_TabCount
        If m_TabCategories(i).Category = CAT_UNCATEGORIZED Then
            count = count + 1
            uncategorizedList = uncategorizedList & "- " & m_TabCategories(i).TabName & vbCrLf
        End If
    Next i
    
    If count > 0 Then
        response = MsgBox("The following tabs were not categorized:" & vbCrLf & vbCrLf & _
                         uncategorizedList & vbCrLf & _
                         "These tabs will be ignored during processing." & vbCrLf & vbCrLf & _
                         "Do you want to proceed?", _
                         vbYesNo + vbQuestion, "Uncategorized Tabs")
        
        If response = vbNo Then
            ' User wants to go back and categorize
            ShowCategorizationDialog
        End If
    End If
End Sub

' Validate that all required categories are assigned
Public Function ValidateCategories() As Boolean
    Dim requiredCategories As Variant
    Dim cat As Variant
    Dim found As Boolean
    Dim i As Long
    Dim missingList As String
    
    ' These categories are required for the tool to work
    requiredCategories = Array(CAT_INPUT_CONTINUING)
    
    For Each cat In requiredCategories
        found = False
        For i = 1 To m_TabCount
            If m_TabCategories(i).Category = cat Then
                found = True
                Exit For
            End If
        Next i
        
        If Not found Then
            missingList = missingList & "- " & cat & vbCrLf
        End If
    Next cat
    
    If missingList <> "" Then
        MsgBox "The following required categories are missing:" & vbCrLf & vbCrLf & _
               missingList & vbCrLf & _
               "Please categorize at least one tab for each required category.", _
               vbCritical, "Missing Required Categories"
        ValidateCategories = False
    Else
        ValidateCategories = True
    End If
End Function

' Get tabs for a specific category
Public Function GetTabsForCategory(categoryName As String) As Collection
    Dim tabs As New Collection
    Dim i As Long
    
    If g_TabCategories.Exists(categoryName) Then
        Set tabs = g_TabCategories(categoryName)
    End If
    
    Set GetTabsForCategory = tabs
End Function

' Get category for a specific tab
Public Function GetCategoryForTab(tabName As String) As String
    Dim i As Long
    
    For i = 1 To m_TabCount
        If m_TabCategories(i).TabName = tabName Then
            GetCategoryForTab = m_TabCategories(i).Category
            Exit Function
        End If
    Next i
    
    GetCategoryForTab = CAT_UNCATEGORIZED
End Function

' Get division name for a segment tab
Public Function GetDivisionName(tabName As String) As String
    Dim i As Long
    
    For i = 1 To m_TabCount
        If m_TabCategories(i).TabName = tabName Then
            GetDivisionName = m_TabCategories(i).DivisionName
            Exit Function
        End If
    Next i
    
    GetDivisionName = ""
End Function
