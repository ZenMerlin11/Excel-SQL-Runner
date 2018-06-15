Attribute VB_Name = "MGlobalConstants"
' ==============================================================================
' File: MGlobalConstants.bas
' Contains global constants for Doc OP Data Validation Tool spreadsheet.
'
' About: Dependencies
' None
'
' About: References
' None
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
' Global Constants
' ------------------------------------------------------------------------------

'General
Public Const lng_EXCEL_MAX_COLUMNS As Long = 16384
Public Const lng_EXCEL_MAX_ROWS As Long = 1048576
Public Const str_DEV_PASSWORD As String = "" 'TODO: Change password

'Sheet Names
Public Const str_CONNECTION As String = "Connection"
Public Const str_SQL_SHEET As String = "SQL" 'sheet where queries are stored
Public Const str_SUMMARY As String = "Summary"
Public Const str_TEST As String = "Test"

'Range Addresses
Public Const str_SERVER_NAME As String = "ServerName"
Public Const str_INSTANCE_NAME As String = "InstanceName"
Public Const str_DATABASE_NAME As String = "DatabaseName"
Public Const str_SQL_AUTHENTICATION As String = "SQLAuthentication"
Public Const str_SQL_LOGIN As String = "SQLLogin"
Public Const str_SQL_QUERY_TOP As String = "FirstQuery"
Public Const str_RUNNING_TIME As String = "RunningTime"
Public Const str_DATA_TAB_FIRST_ROW As String = "A6"
Public Const str_ACCOUNT_ID As String = "AccountID"
Public Const str_CUSTOMER_NAME As String = "CustomerName"

'Significant Row and Column Numbers
Public Const lng_DATA_TAB_FIRST_ROW As Long = 6 '1st row number on data tabs
Public Const lng_SQL_QUERY_TOP_ROW As Long = 3 '1st row of queries on SQL tab

'Authentification Types
Public Const str_WINDOWS_AUTH As String = "Windows"
Public Const str_SQL_SERVER_AUTH As String = "SQL Server"
