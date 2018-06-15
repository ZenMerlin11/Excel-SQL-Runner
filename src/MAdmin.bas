Attribute VB_Name = "MAdmin"
' ==============================================================================
' File: MAdmin.bas
' Controls developer access to hidden worksheets and other workbook cleanup
' utilities.
'
' About: Dependencies
' MGlobalConstants.bas
'
' About: References
' Microsoft ActiveX Data Objects 2.8 Library
'
' About: Compatibility
' Excel 2013, 2016
'
' About: License
' This file is licensed under the MIT license.
'
' About: Author
' Jason Boyll
'
' jason.boyll@gmail.com
' ==============================================================================

'@Folder("Modules")

' ------------------------------------------------------------------------------
' Option Statements
' ------------------------------------------------------------------------------

Option Explicit


' ------------------------------------------------------------------------------
' Subs and Functions
' ------------------------------------------------------------------------------

' Sub: ToggleDevMode
'
' Toggles developer mode on or off. Checks visibility of SQL tab to determine
' developer mode (visible = on, hidden = off). Queries user for password if
' enabling developer mode.
Public Sub ToggleDevMode()

    Dim strPassword As String
    
    Application.ScreenUpdating = False 'turn off screen updates

    If ThisWorkbook.Sheets(str_SQL_SHEET).Visible <> xlSheetVisible Then
        If Not EnableDevMode(InputBox("Enter Developer Password: ")) Then
            MsgBox "Incorrect Password"
        Else
            MsgBox "Dev Mode Enabled"
            
            'Activate SQL Sheet and select first cell
            With ThisWorkbook.Sheets(str_SQL_SHEET)
                .Activate
                .Range("A1").Select
            End With
        End If
    Else
        DisableDevMode
        MsgBox "Dev Mode Disabled"
        
        'Activate Connection sheet and select first cell
        With ThisWorkbook.Sheets(str_CONNECTION)
            .Activate
            .Range("A1").Select
        End With
        
    End If
    
    Application.ScreenUpdating = True 'turn on screen updates
    
End Sub


' Sub: EnableDevMode
'
' Unhides developer worksheets (SQL, Test, and all other output sheets).
' Requires password.
Public Function EnableDevMode(ByVal password As String) As Boolean
    
    Dim wks As Worksheet
    Dim i As Long
    
    EnableDevMode = False 'default to false
    
    'Check password to allow dev access
    If password = str_DEV_PASSWORD Then
        For i = 1 To ThisWorkbook.Worksheets.Count
            Set wks = ThisWorkbook.Sheets(i)
            wks.Visible = xlSheetVisible
        Next i
        EnableDevMode = True 'return true if access granted
    End If
    
End Function


' Sub: DisableDevMode
'
' Hides developer worksheets (SQL, Test, and all other output sheets except
' Connection).
Public Sub DisableDevMode()

    Dim wks As Worksheet
    Dim i As Long
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        If wks.Name <> str_CONNECTION Then
            wks.Visible = xlSheetVeryHidden
        End If
    Next i
        
End Sub


' Sub: CleanUpWorksheets
'
' Cleans up worksheets after data population.
Public Sub CleanUpWorksheets()
    
    Dim rngPrevSel As Range
    Dim wksPrevSel As Worksheet
    Dim wks As Worksheet
    
    Application.ScreenUpdating = False 'disable screen updates while running
    
        
    'Reset selection
    For Each wks In ThisWorkbook.Worksheets
        wks.Activate
        wks.Range("A1").Select
    Next wks
    
    'Turn on and select summary tab
    With ThisWorkbook.Sheets(str_SUMMARY)
        .Visible = xlSheetVisible
        .Activate
        .Range("A1").Select
    End With
    
    'Clear SQL password
    With ThisWorkbook.Sheets(str_CONNECTION)
        .txtPassword.Text = vbNullString
    End With
    
    Application.ScreenUpdating = True 'turn screen updates back on
    
End Sub


' Sub: Reset
'
' Iterates through each worksheet and clears all data sheets from row 6 down.
' Resets visibility of each data sheet to xlSheetVeryHidden.
Public Sub Reset()
    
    Dim wks As Worksheet
    Dim i As Long
    
    Set wks = ThisWorkbook.Sheets(str_CONNECTION)
    wks.Range(str_RUNNING_TIME).Value = vbNullString
    
    'Clear data from worksheets and reset visibility
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        If Not (wks.Name = str_CONNECTION Or wks.Name = str_SQL_SHEET Or _
            wks.Name = str_SUMMARY Or wks.Name = str_TEST) Then
            wks.Rows(lng_DATA_TAB_FIRST_ROW & ":" & wks.Rows.Count).Clear
            wks.Visible = xlSheetVeryHidden
        End If
    Next i
    
    'Set Summary, SQL, and Test sheets to hidden
    With ThisWorkbook
        Sheets(str_SUMMARY).Visible = xlVeryHidden
        Sheets(str_SQL_SHEET).Visible = xlVeryHidden
        Sheets(str_TEST).Visible = xlVeryHidden
    End With
    
    
End Sub


