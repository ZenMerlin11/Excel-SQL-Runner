Attribute VB_Name = "MTestSQLRunner"
' ==============================================================================
' File: MTestSQLRunner.bas
' Test module for MSQLRunner.bas
'
' About: Dependencies
' Rubberduck Add-in (Rubberduckvba.com), MGlobalConstants.bas,
' MSQLRunner.bas
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
Public Sub TestCreateConnectionString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim strInstance As String
    Dim strDatabase As String
    Dim strAuthentication As String
    Dim strLogin As String
    Dim strPassword As String
    Dim strWinExpected As String
    Dim strSQLExpected As String
    Dim strWinResult As String
    Dim strSQLResult As String
        
    
    With ThisWorkbook.Sheets(str_CONNECTION)
        'Setup for windows authentication test
        .Range(str_SQL_AUTHENTICATION).Value = str_WINDOWS_AUTH
        
        'Retrieve database login info from connection worksheet
        strInstance = .Range(str_SERVER_NAME) & "\" & .Range(str_INSTANCE_NAME)
        strDatabase = .Range(str_DATABASE_NAME)
        strAuthentication = .Range(str_SQL_AUTHENTICATION)
        strLogin = .Range(str_SQL_LOGIN)
        strPassword = .txtPassword.Text
    End With
    
    'Define expected value for windows authentication
    strWinExpected = _
        "Driver={SQL Server};" & _
        "Server=" & strInstance & ";" & _
        "Database=" & strDatabase & ";" & _
        "Trusted_Connection=Yes;"
    
    strSQLExpected = _
        "Driver={SQL Server};" & _
        "Server=" & strInstance & ";" & _
        "Database=" & strDatabase & ";" & _
        "Uid=" & strLogin & ";" & _
        "Pwd=" & strPassword & ";"

    'Act:
    'Windows Authentication Result
    strWinResult = CreateConnectionString()
    
    'SQL Server Authentication Result
    With ThisWorkbook.Sheets(str_CONNECTION)
        'Set authentication input to SQL Server
        .Range(str_SQL_AUTHENTICATION).Value = str_SQL_SERVER_AUTH
    End With
    
    strSQLResult = CreateConnectionString()

    'Assert:
    Assert.AreEqual strWinExpected, strWinResult, _
        "Windows authentication case failed."
    Assert.AreEqual strSQLExpected, strSQLResult, _
        "SQL Server authentication case failed."

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub TestProcessQuery()
    On Error GoTo TestFail
    
    'Arrange:
    Dim strExpected As String
    Dim strResult1 As String
    Dim strResult2 As String
    Dim strResult3 As String
    
    strExpected = "Hello World"
    
    'Act:
    strResult1 = ProcessQuery(testWks.Range("B10"))
    strResult2 = ProcessQuery(testWks.Range("B11"))
    strResult3 = ProcessQuery(testWks.Range("B12"))

    'Assert:
    'Check that all three results equal the expected value
    Assert.AreEqual strExpected, strResult1, "Case 1 failed"
    Assert.AreEqual strExpected, strResult2, "Case 2 failed"
    Assert.AreEqual strExpected, strResult3, "Case 3 failed"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & _
        Err.Description
End Sub


'@TestMethod
Public Sub TestRecordsFound()
    On Error GoTo TestFail
    
    Dim rng As Range
    Dim i As Long
    
    'Case no records exist
    Set rng = testWks.Range(str_DATA_TAB_FIRST_ROW)
    
    For i = 1 To 5
        'Create header
        rng.Offset(-1, i - 1).Value = "Col Header " & i
    Next i
    
    Assert.IsFalse RecordsFound(testWks), "Case 'no records exist' failed."
    
    'Case records exist
    'Add a dummy value to first row
    rng.Value = "A"
    
    Assert.IsTrue RecordsFound(testWks), "Case 'records exist' failed."
    
    'Cleanup
    testWks.Range("A5:E6").Clear
    

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & _
        Err.Description
End Sub

