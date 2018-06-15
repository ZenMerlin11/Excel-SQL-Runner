Attribute VB_Name = "MTestWorkSheetFunctions"
' ==============================================================================
' File: MTestWorkSheetFunctions.bas
' Test module for MTestWorkSheetFunctions.bas
'
' About: Dependencies
' Rubberduck Add-in (Rubberduckvba.com), MGlobalConstants.bas,
' MWorkSheetFunctions.bas
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
Public Sub TestCountU() 'TODO Rename test
    On Error GoTo TestFail
    
    'Arrange:
    Dim rng As Range
    Dim rng1 As Range
    Dim rng2 As Range
    Dim rng3 As Range
    
    Dim expected1 As Long
    Dim expected2 As Long
    Dim expected3 As Long
    
    Dim result1 As Long
    Dim result2 As Long
    Dim result3 As Long
            
    Dim i As Long
    Dim j As Long
        
    testWks.Activate
        
    Set rng1 = testWks.Range("A15:A54") '40R
    Set rng2 = testWks.Range("B15:AA15") '26C
    Set rng3 = testWks.Range("B16:F25") '10R X 5C
    
    'Populate ranges with values
    Set rng = testWks.Range("A15")
    For i = 1 To 5
        rng.Cells(i, 1).Value = "alpha"
    Next i
    
    For i = 6 To 10
        rng.Cells(i, 1).Value = "beta"
    Next i
    
    For i = 11 To 20
        rng.Cells(i, 1).Value = 5
    Next i
    
    expected1 = 3
    
    Set rng = testWks.Range("B15")
    For i = 1 To 10
        rng.Cells(1, i).Value = "charlie"
    Next i
    
    For i = 11 To 15
        rng.Cells(1, i).Value = 6
    Next i
    
    expected2 = 2
    
    Set rng = testWks.Range("B16")
    For i = 1 To 2
        For j = 1 To 2
            rng.Cells(i, j).Value = 1
        Next j
    Next i
    
    For i = 3 To 4
        For j = 3 To 4
            rng.Cells(i, j).Value = 2
        Next j
    Next i
    
    expected3 = 2
    
    'Act:
    result1 = CountU(rng1)
    result2 = CountU(rng2)
    result3 = CountU(rng3)
    
    'Assert:
    Assert.AreEqual expected1, result1, "Case 1 failed."
    Assert.AreEqual expected2, result2, "Case 2 failed."
    Assert.AreEqual expected3, result3, "Case 3 failed."
    
    'Clean up test worksheet:
    testWks.Range("A15:P34").Delete

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & _
        Err.Description
End Sub
