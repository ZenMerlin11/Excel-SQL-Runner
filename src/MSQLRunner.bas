Attribute VB_Name = "MSQLRunner"
' ==============================================================================
' File: MSQLRunner.bas
' Runs user defined queries on a given database and populates corresponding
' worksheets with the results.
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

' Sub: StartExtraction
'
' Runs user defined queries on database and populates corresponding sheets with
' retrieved data. Main entry point for data extraction.
Public Sub StartExtraction()
    
    'Time macro execution
    Dim dblStartTime As Double
    Dim strTimeElapsed As String
    
    'Remember time when macro starts
    dblStartTime = Timer
    
    'Database declaration
    Dim dbConnection As ADODB.Connection
    
    'Open connection to DocOP database
    Set dbConnection = New ADODB.Connection
    
    'Run queries if connection successful
    If OpenDatabase(dbConnection) Then
        RunQueries dbConnection
        
        'Close connection
        dbConnection.Close
        Set dbConnection = Nothing
    
        'Clean up worksheets
        CleanUpWorksheets
    End If
    
    'Notify user of time elapsed
    strTimeElapsed = Format$((Timer - dblStartTime) / 86400, "hh:mm:ss")
    ThisWorkbook.Sheets(str_CONNECTION).Range(str_RUNNING_TIME).Value = _
        strTimeElapsed

End Sub


' Sub: OpenDatabase
'
' Attempts to open the requested database. Returns true if successful.
Public Function OpenDatabase(ByVal dbConnection As ADODB.Connection) As Boolean

    Dim strError As String
    Dim strConnection As String
        
    On Error GoTo AdoError
    
    OpenDatabase = True 'the default return value is True (success)
    
    strConnection = CreateConnectionString()
    
    If strConnection <> vbNullString Then
        dbConnection.Open strConnection
    Else
        OpenDatabase = False
    End If
    
    'Indefinite execute timeout
    If OpenDatabase Then
        dbConnection.CommandTimeout = 0
    End If

Done:
    Exit Function

AdoError:
    OpenDatabase = False
    On Error Resume Next
    
    Dim strInstance As String
    
    With ThisWorkbook.Sheets(str_CONNECTION)
        strInstance = .Range(str_SERVER_NAME) & "\" & .Range(str_INSTANCE_NAME)
    End With
    
    strError = "An error occurred connecting to " & strInstance & vbCrLf & _
        "Please make sure that the server name is correct and that the SQL" & _
        "Service is started."
    
    'Display the connection error
    MsgBox strError
    
    'Clean up gracefully without risking infinite loop in error handler
    On Error GoTo 0
    Set dbConnection = Nothing
    
End Function


' Sub: CreateConnectionString
'
' Returns the database connection string based on input values entered in the
' connection worksheet.
Public Function CreateConnectionString() As String
    
    Dim strInstance As String
    Dim strDatabase As String
    Dim strAuthentication As String
    Dim strLogin As String
    Dim strPassword As String
    
    'Retrieve database login info from connection worksheet
    With ThisWorkbook.Sheets(str_CONNECTION)
        strInstance = .Range(str_SERVER_NAME) & "\" & .Range(str_INSTANCE_NAME)
        strDatabase = .Range(str_DATABASE_NAME)
        strAuthentication = .Range(str_SQL_AUTHENTICATION)
        strLogin = .Range(str_SQL_LOGIN)
        strPassword = .txtPassword.Text
    End With
    
    'Create connection string from database login info above
    Select Case strAuthentication
        'Trusted connection (Windows security)
        Case str_WINDOWS_AUTH
            CreateConnectionString = _
                "Driver={SQL Server};" & _
                "Server=" & strInstance & ";" & _
                "Database=" & strDatabase & ";" & _
                "Trusted_Connection=Yes;"
        
        'Standard security
        Case str_SQL_SERVER_AUTH
            CreateConnectionString = _
                "Driver={SQL Server};" & _
                "Server=" & strInstance & ";" & _
                "Database=" & strDatabase & ";" & _
                "Uid=" & strLogin & ";" & _
                "Pwd=" & strPassword & ";"
        
        'Invalid authentification type specified
        Case Else
            CreateConnectionString = vbNullString
    End Select
    
End Function


' Sub: RunQueries
'
' Reads through and executes all consecutive queries stored in SQL tab on the
' specified database.
Public Sub RunQueries(ByVal dbConnection As ADODB.Connection)
    
    Dim rng As Range
    Dim i As Long
    
    Set rng = ThisWorkbook.Sheets(str_SQL_SHEET).Range(str_SQL_QUERY_TOP)
    
    i = 1
    Do While rng.Cells(i, 1).Value <> vbNullString
        RunQuery rng.Cells(i, 1), dbConnection
        i = i + 1
        
        'Exit loop if max row reached
        If i >= (lng_EXCEL_MAX_ROWS - lng_SQL_QUERY_TOP_ROW + 1) Then
            Exit Do
        End If
    Loop
    
End Sub


' Sub: RunQuery
'
' Executes the query from the given range on the specified database and copies
' records found to the corresponding output worksheet. If data is written to a
' worksheet, it is made visible. Worksheets that are marked as hidden on the
' SQL sheet remain hidden regardless of whether they were written to.
Public Sub RunQuery(ByVal query As Range, _
    ByVal dbConnection As ADODB.Connection)
    
    Dim strQuery As String
    Dim strSheetName As String
    Dim blnSheetHidden As Boolean
    Dim rst As ADODB.Recordset
    Dim wks As Worksheet
    
    Set rst = New ADODB.Recordset
    
    'Get query and target output sheet parameters
    strQuery = ProcessQuery(query)
    strSheetName = query.Cells(1, 2).Value
    blnSheetHidden = query.Cells(1, 3).Value
    
    'Get recordset and copy to target sheet if records found
    rst.Open strQuery, dbConnection
    If rst.State = adStateOpen Then
        Set wks = ThisWorkbook.Sheets(strSheetName)
        
        wks.Range(str_DATA_TAB_FIRST_ROW).CopyFromRecordset rst
        
        If Not blnSheetHidden And RecordsFound(wks) Then
            wks.Visible = xlSheetVisible
        End If
        
        rst.Close 'close recordset when done
    End If

End Sub


' Sub: ProcessQuery
'
' Takes range where query is located and looks up cells to the right for
' find/replace strings to change in query. If adjacent cells are empty, the
' original query is returned unaltered.
Public Function ProcessQuery(ByVal query As Range) As String
    
    Dim rng As Range
    Dim strFind As String
    Dim strReplace As String
    Dim strQuery As String
    Dim i As Long
        
    strQuery = query.Cells(1, 1).Value
    Set rng = query.Cells(1, 4) 'get range of beginning find/replace
    
    'Process find/replace text strings adjacent to the Query column
    i = 1
    Do While rng.Cells(1, i).Value <> vbNullString
        strFind = rng.Cells(1, i).Value
        strReplace = rng.Cells(1, i + 1).Value
        strQuery = Replace(strQuery, strFind, strReplace)
        
        i = i + 2 'skip to the next find string
        
        'Exit loop if max column reached
        If i >= lng_EXCEL_MAX_COLUMNS Then
            Exit Do
        End If
    Loop
    
    ProcessQuery = strQuery
    
End Function


' Sub: RecordsFound
'
' Takes a data sheet and checks the first row for data. If row is empty,
' the function returns false indicating that no records were found.
Public Function RecordsFound(ByVal wks As Worksheet) As Boolean
    
    Dim rngRow As Range
    Dim rngHeader As Range
    Dim i As Long
        
    RecordsFound = False 'default to no records found
    
    Set rngRow = wks.Range(str_DATA_TAB_FIRST_ROW)
    Set rngHeader = rngRow.Offset(-1, 0)
    
    i = 1
    
    'Return true if any cells along first row contain other than a null string
    Do While rngHeader.Cells(1, i).Value <> vbNullString
        If rngRow.Cells(1, i).Value <> vbNullString Then
            RecordsFound = True
        End If
        
        i = i + 1
        
        'Exit loop if max column reached
        If i >= lng_EXCEL_MAX_COLUMNS Then
            Exit Do
        End If
    Loop
    
End Function

