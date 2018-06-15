Attribute VB_Name = "MWorkSheetFunctions"
' ==============================================================================
' File: MWorkSheetFunctions.bas
' Contains user defined worksheet functions used in this workbook
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

' Sub: CountU
'
' Worksheet function that takes a range and returns the number of unique values
' found in the range as a long.
Public Function CountU(ByVal rng As Range) As Long

    Dim colUniqueElements As Collection
    Dim blnFound As Boolean
    Dim varElement As Variant
    
    Set colUniqueElements = New Collection
    
    Dim i As Long
    Dim j As Long
    
    For i = 1 To rng.Rows.Count
        For j = 1 To rng.Columns.Count
            varElement = rng.Cells(i, j).Value
            If Not FoundInCollection(colUniqueElements, varElement) _
                And varElement <> vbNullString Then
                colUniqueElements.Add varElement
            End If
        Next j
    Next i
    
    CountU = colUniqueElements.Count
    
End Function


' Sub: FoundInCollection
'
' Helper function for CountU. Returns true if match is found in col.
Private Function FoundInCollection(ByVal col As Collection, _
    ByVal match As Variant) As Boolean
    
    Dim varElement As Variant
    
    FoundInCollection = False 'default to not found
    
    If col.Count <> 0 Then
        For Each varElement In col
            If varElement = match Then
                FoundInCollection = True
                Exit For
            End If
        Next varElement
    End If

End Function
