Attribute VB_Name = "MTestAdmin"
' ==============================================================================
' File: MTestAdmin.bas
' Test module for MAdmin.bas
'
' About: Dependencies
' Rubberduck Add-in (Rubberduckvba.com), MGlobalConstants.bas,
' MAdmin.bas
'
' About: References
' None
'
' About: Compatibility
' Excel 2013 - 2016
'
' About: License
' This file is licensed under the MIT license.
'
' About: Author
' Jason Boyll
'
' jason.boyll@gmail.com
' ==============================================================================

'@TestModule
'@Folder("Tests")


' ------------------------------------------------------------------------------
' Option Statements
' ------------------------------------------------------------------------------

Option Explicit
Option Private Module


' ------------------------------------------------------------------------------
' Module Level Constants and Variables
' ------------------------------------------------------------------------------

' Test objects
Private Assert As Object
Private Fakes As Object
Private testWks As Worksheet


' ------------------------------------------------------------------------------
' Subs and Functions
' ------------------------------------------------------------------------------

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub


'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
    Set testWks = ThisWorkbook.Sheets("test")
    Application.ScreenUpdating = False
End Sub


'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
    Set testWks = Nothing
    Application.ScreenUpdating = True
End Sub


'@TestMethod
Public Sub TestEnableDevMode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim blnResult As Boolean
    Dim wks As Worksheet
    Dim i As Long
            
    'Disable Dev Mode by hiding all sheets except Connection
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        If wks.Name <> str_CONNECTION Then
            wks.Visible = xlSheetVeryHidden
        End If
    Next i
    
    'Assert Cases:
    'Test password fail condition
    Assert.IsFalse EnableDevMode("Wrong Password"), _
        "Should return false when wrong password supplied."
    
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        If wks.Name <> str_CONNECTION Then
            Assert.AreEqual wks.Visible, xlSheetVeryHidden, _
                "All sheets but Connection should be hidden."
        End If
    Next i
    
    'Test access granted condition
    Assert.IsTrue EnableDevMode(str_DEV_PASSWORD), _
        "Should return true when correct password supplied."
        
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        Assert.AreEqual wks.Visible, xlSheetVisible, _
            "All worksheets should be visible in dev mode."
    Next i
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & _
        Err.Description
End Sub


'@TestMethod
Public Sub TestDisableDevMode()
    On Error GoTo TestFail
    
    'Arrange:
    Dim blnResult As Boolean
    Dim wks As Worksheet
    Dim i As Long
    
    'Enable dev mode by showing all sheets
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        wks.Visible = xlSheetVisible
    Next i
            
    'Act:
    DisableDevMode
        
    'Assert:
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        If wks.Name <> str_CONNECTION Then
            Assert.AreEqual wks.Visible, xlSheetVeryHidden, _
                "All sheets except ""Connection"" should be hidden."
        End If
    Next i

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & _
        Err.Description
End Sub


'@TestMethod
Public Sub TestReset()
    On Error GoTo TestFail
    
    'Arrange:
    Dim i As Long
    Dim wks As Worksheet
    Const dblCellsNotEmpty As Double = 0
    
    'Act:
    Reset

    'Assert:
    'Check data worksheets are cleared and hidden
    For i = 1 To ThisWorkbook.Worksheets.Count
        Set wks = ThisWorkbook.Sheets(i)
        If Not (wks.Name = str_CONNECTION Or wks.Name = str_SQL_SHEET Or _
            wks.Name = str_SUMMARY Or wks.Name = str_TEST) Then
            
            'Check that rows after headers are blank
            Assert.AreEqual _
                WorksheetFunction.CountA(wks.Range("A6:XFD1048576")), _
                dblCellsNotEmpty, "Rows 6 down are not all blank on sheet " & _
                wks.Name
            
            'Check that sheet is hidden
            Assert.AreEqual wks.Visible, xlSheetVeryHidden, _
                "Sheet " & wks.Name & " should be hidden (xlSheetVeryHidden)"
        End If
    Next i
    
    'Check that running time cell is cleared
    Set wks = ThisWorkbook.Sheets(str_CONNECTION)
    Assert.AreEqual wks.Range(str_RUNNING_TIME).Value, vbNullString, _
        "Running time was not cleared."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & _
        Err.Description
End Sub

