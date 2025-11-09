Option Explicit

' ============================================================================
' MODULE: ModTabCategorization
' PURPOSE: Handle tab categorization and validation
' DESCRIPTION: Provides functionality to categorize worksheets into predefined
'              categories for proper processing
' ============================================================================

' Category constants
Public Const CAT_SEGMENT = "TGK Segment Tabs"
Public Const CAT_DISCONTINUED = "Discontinued Ops Tab"
Public Const CAT_INPUT_CONTINUING = "TGK Input Continuing Operations Tab"
Public Const CAT_JOURNALS_CONTINUING = "TGK Journals Continuing Tab"
Public Const CAT_CONSOLE_CONTINUING = "TGK Consol Continuing Tab"
Public Const CAT_BS = "TGK BS Tab"
Public Const CAT_IS = "TGK IS Tab"
Public Const CAT_PULL_WORKINGS = "Paul workings"
Public Const CAT_TRIAL_BALANCE = "Trial Balance"
Public Const CAT_UNCATEGORIZED = "Uncategorized"

' Structure to hold tab categorization
Private Type TabCategory
    tabName As String
    Category As String
    divisionName As String ' Used for segment tabs
End Type

Private m_TabCategories() As TabCategory
Private m_TabCount As Long

' Initialize the categorization system
Public Function CategorizeTabs(tabList As Collection) As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim userForm As Object
    
    ' Initialize array
    m_TabCount = tabList.count
    ReDim m_TabCategories(1 To m_TabCount)
    
    ' Populate tab names
    For i = 1 To tabList.count
        m_TabCategories(i).tabName = tabList(i)
        m_TabCategories(i).Category = CAT_UNCATEGORIZED
        m_TabCategories(i).divisionName = ""
    Next i
    
    ' Show categorization interface
    If Not ShowCategorizationDialog() Then
        CategorizeTabs = False
        Exit Function
    End If
    
    ' Store categorization in global dictionary
    ' Note: Requires Microsoft Scripting Runtime library
    On Error Resume Next
    Set g_TabCategories = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Error: Cannot create Scripting.Dictionary object." & vbCrLf & vbCrLf & _
               "Please ensure Microsoft Scripting Runtime is available:" & vbCrLf & _
               "1. Open VBA Editor (Alt+F11)" & vbCrLf & _
               "2. Go to Tools > References" & vbCrLf & _
               "3. Check 'Microsoft Scripting Runtime'" & vbCrLf & _
               "4. Click OK and try again", vbCritical, "Missing Library"
        CategorizeTabs = False
        Exit Function
    End If
    On Error GoTo ErrorHandler
    
    For i = 1 To m_TabCount
        If Not g_TabCategories.Exists(m_TabCategories(i).Category) Then
            g_TabCategories.Add m_TabCategories(i).Category, New Collection
        End If
        
        Dim tabInfo As Object
        On Error Resume Next
        Set tabInfo = CreateObject("Scripting.Dictionary")
        If Err.Number <> 0 Then
            On Error GoTo 0
            MsgBox "Error: Cannot create Scripting.Dictionary object." & vbCrLf & vbCrLf & _
                   "Please ensure Microsoft Scripting Runtime is available:" & vbCrLf & _
                   "1. Open VBA Editor (Alt+F11)" & vbCrLf & _
                   "2. Go to Tools > References" & vbCrLf & _
                   "3. Check 'Microsoft Scripting Runtime'" & vbCrLf & _
                   "4. Click OK and try again", vbCritical, "Missing Library"
            CategorizeTabs = False
            Exit Function
        End If
        On Error GoTo ErrorHandler
        
        tabInfo("TabName") = m_TabCategories(i).tabName
        tabInfo("DivisionName") = m_TabCategories(i).divisionName
        
        g_TabCategories(m_TabCategories(i).Category).Add tabInfo
    Next i
    
    CategorizeTabs = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in tab categorization: " & Err.Description, vbCritical
    CategorizeTabs = False
End Function

' Show dialog for categorizing tabs using pop-up dialogs
Private Function ShowCategorizationDialog() As Boolean
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim response As VbMsgBoxResult
    Dim categoryChoice As String
    Dim divisionName As String
    Dim categoryList As String
    Dim categoryNumber As Integer
    Dim continueLoop As Boolean
    Dim validationPassed As Boolean
    Dim startOver As Boolean
    
    ' Show instructions
    MsgBox "Tab Categorization - Pop-up Mode" & vbCrLf & vbCrLf & _
           "You will now categorize each tab using pop-up dialogs." & vbCrLf & vbCrLf & _
           "For each tab, you'll select a category by entering a number:" & vbCrLf & vbCrLf & _
           "1 = " & CAT_SEGMENT & " (multiple allowed)" & vbCrLf & _
           "2 = " & CAT_DISCONTINUED & " (single only *)" & vbCrLf & _
           "3 = " & CAT_INPUT_CONTINUING & " (single only * - REQUIRED)" & vbCrLf & _
           "4 = " & CAT_JOURNALS_CONTINUING & " (single only *)" & vbCrLf & _
           "5 = " & CAT_CONSOLE_CONTINUING & " (single only *)" & vbCrLf & _
           "6 = " & CAT_BS & " (single only *)" & vbCrLf & _
           "7 = " & CAT_IS & " (single only *)" & vbCrLf & _
           "8 = " & CAT_PULL_WORKINGS & " (multiple allowed)" & vbCrLf & _
           "9 = " & CAT_TRIAL_BALANCE & " (single only *)" & vbCrLf & _
           "10 = " & CAT_UNCATEGORIZED & " (skip this tab)", _
           vbInformation, "Categorization Instructions"
    
    ' Build category list for reference
    categoryList = "1 = " & CAT_SEGMENT & vbCrLf & _
                   "2 = " & CAT_DISCONTINUED & vbCrLf & _
                   "3 = " & CAT_INPUT_CONTINUING & vbCrLf & _
                   "4 = " & CAT_JOURNALS_CONTINUING & vbCrLf & _
                   "5 = " & CAT_CONSOLE_CONTINUING & vbCrLf & _
                   "6 = " & CAT_BS & vbCrLf & _
                   "7 = " & CAT_IS & vbCrLf & _
                   "8 = " & CAT_PULL_WORKINGS & vbCrLf & _
                   "9 = " & CAT_TRIAL_BALANCE & vbCrLf & _
                   "10 = " & CAT_UNCATEGORIZED
    
    ' Main categorization loop with retry capability
    validationPassed = False
    Do While Not validationPassed
        ' Loop through each tab and get category
        For i = 1 To m_TabCount
            continueLoop = True
            
            Do While continueLoop
                ' Prompt for category
                categoryChoice = InputBox( _
                    "Tab " & i & " of " & m_TabCount & vbCrLf & vbCrLf & _
                    "Tab Name: " & m_TabCategories(i).tabName & vbCrLf & vbCrLf & _
                    "Select a category (enter number 1-10):" & vbCrLf & vbCrLf & _
                    categoryList, _
                    "Categorize Tab", _
                    "3")
                
                ' Check if user cancelled
                If categoryChoice = "" Then
                    response = MsgBox("Do you want to cancel the categorization process?", _
                                     vbYesNo + vbQuestion, "Cancel Categorization")
                    If response = vbYes Then
                        ShowCategorizationDialog = False
                        Exit Function
                    Else
                        ' Continue with current tab - stay in loop to show InputBox again
                        continueLoop = True
                    End If
                Else
                    ' Validate input
                    If IsNumeric(categoryChoice) Then
                        categoryNumber = CInt(categoryChoice)
                        
                        Select Case categoryNumber
                            Case 1
                                m_TabCategories(i).Category = CAT_SEGMENT
                                ' Prompt for division name
                                divisionName = InputBox( _
                                    "Enter the division name for this segment tab:" & vbCrLf & vbCrLf & _
                                    "Tab: " & m_TabCategories(i).tabName & vbCrLf & vbCrLf & _
                                    "Examples: UK Division, Properties Division, BIH division, etc.", _
                                    "Enter Division Name", _
                                    "")
                                
                                ' If division name is empty, prompt again or use default
                                If Trim(divisionName) = "" Then
                                    divisionName = "Division_" & i
                                End If
                                m_TabCategories(i).divisionName = Trim(divisionName)
                                continueLoop = False
                            Case 2
                                m_TabCategories(i).Category = CAT_DISCONTINUED
                                continueLoop = False
                            Case 3
                                m_TabCategories(i).Category = CAT_INPUT_CONTINUING
                                continueLoop = False
                            Case 4
                                m_TabCategories(i).Category = CAT_JOURNALS_CONTINUING
                                continueLoop = False
                            Case 5
                                m_TabCategories(i).Category = CAT_CONSOLE_CONTINUING
                                continueLoop = False
                            Case 6
                                m_TabCategories(i).Category = CAT_BS
                                continueLoop = False
                            Case 7
                                m_TabCategories(i).Category = CAT_IS
                                continueLoop = False
                            Case 8
                                m_TabCategories(i).Category = CAT_PULL_WORKINGS
                                continueLoop = False
                            Case 9
                                m_TabCategories(i).Category = CAT_TRIAL_BALANCE
                                continueLoop = False
                            Case 10
                                m_TabCategories(i).Category = CAT_UNCATEGORIZED
                                continueLoop = False
                            Case Else
                                MsgBox "Invalid number. Please enter a number between 1 and 10.", vbExclamation
                                continueLoop = True
                        End Select
                    Else
                        MsgBox "Invalid input. Please enter a number between 1 and 10.", vbExclamation
                        continueLoop = True
                    End If
                End If
            Loop
        Next i
        
        ' Validate single-tab categories
        If ValidateSingleTabCategories() Then
            ' Show uncategorized tabs and check if user wants to proceed
            If ShowUncategorizedTabs() Then
                validationPassed = True
            Else
                ' User wants to restart - reset and continue loop
                response = MsgBox("Do you want to start over with categorization?", _
                                 vbYesNo + vbQuestion, "Restart Categorization")
                If response = vbYes Then
                    ' Reset and try again
                    For i = 1 To m_TabCount
                        m_TabCategories(i).Category = CAT_UNCATEGORIZED
                        m_TabCategories(i).divisionName = ""
                    Next i
                    ' Loop will continue
                Else
                    ShowCategorizationDialog = False
                    Exit Function
                End If
            End If
        Else
            response = MsgBox("Validation failed. Would you like to start over?", _
                             vbYesNo + vbQuestion, "Validation Error")
            If response = vbYes Then
                ' Reset and try again
                For i = 1 To m_TabCount
                    m_TabCategories(i).Category = CAT_UNCATEGORIZED
                    m_TabCategories(i).divisionName = ""
                Next i
                ' Loop will continue
            Else
                ShowCategorizationDialog = False
                Exit Function
            End If
        End If
    Loop
    
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
                            CAT_CONSOLE_CONTINUING, CAT_BS, CAT_IS, CAT_TRIAL_BALANCE)
    
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

' Show uncategorized tabs to user and return whether to continue
Private Function ShowUncategorizedTabs() As Boolean
    Dim i As Long
    Dim uncategorizedList As String
    Dim count As Long
    Dim response As VbMsgBoxResult
    
    uncategorizedList = ""
    count = 0
    
    For i = 1 To m_TabCount
        If m_TabCategories(i).Category = CAT_UNCATEGORIZED Then
            count = count + 1
            uncategorizedList = uncategorizedList & "- " & m_TabCategories(i).tabName & vbCrLf
        End If
    Next i
    
    If count > 0 Then
        response = MsgBox("The following tabs were not categorized:" & vbCrLf & vbCrLf & _
                         uncategorizedList & vbCrLf & _
                         "These tabs will be ignored during processing." & vbCrLf & vbCrLf & _
                         "Do you want to proceed?", _
                         vbYesNo + vbQuestion, "Uncategorized Tabs")
        
        If response = vbNo Then
            ' User wants to restart categorization
            ShowUncategorizedTabs = False
            Exit Function
        End If
    End If
    
    ShowUncategorizedTabs = True
End Function

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
        If m_TabCategories(i).tabName = tabName Then
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
        If m_TabCategories(i).tabName = tabName Then
            GetDivisionName = m_TabCategories(i).divisionName
            Exit Function
        End If
    Next i
    
    GetDivisionName = ""
End Function
