# VBA-Functions


'Takes a string value and looks for that value in the range row provided. _
 The return value is the letter of the column in which the string is found. _
 The input for lookupSheet is a 'STRING'.
Private Function StringToColumn(lookupValue As String, lookupRow As Integer, lookupSheet As String) As String

    Dim sheetRange As String
    sheetRange = lookupSheet

    Dim iCell As Range
    
    Dim cellNumber As Variant
    
    Dim defineRng As Range
    Set defineRng = Worksheets(sheetRange).Rows(lookupRow).Find(lookupValue)
    
    With Worksheets(sheetRange)

        'Check if value is in range
        If Not defineRng Is Nothing Then
            cellNumber = defineRng.Column
        Else
            cellNumber = CVErr(xlErrNA)
        End If
        
        'If value is in range then find the column number that the value is in, _
         then convert column number to the column letter
        If Not IsError(cellNumber) Then
            For Each iCell In Intersect(.Columns(cellNumber), .UsedRange)
                Debug.Print iCell.Address
                StringToColumn = Split(Cells(lookupRow, cellNumber).Address, "$")(1)
            Next iCell
        Else
            'Error Message
            msgbox "The string '" & lookupValue & "' was not found on sheet '" & sheetRange & "'"
        End If
        
    End With
    
End Function

'Takes a string value and looks for that value in the range row provided. _
 The return value is the letter of the column in which the string is found. _
 The input for lookupSheet is a 'WORKSHEET'.
Private Function StringToColumn(lookupValue As String, lookupRow As Integer, lookupSheet As Worksheet) As String

    Dim sheetRange As Worksheet
    Set sheetRange = lookupSheet

    Dim iCell As Range
    
    Dim cellNumber As Variant
    
    Dim defineRng As Range
    Set defineRng = sheetRange.Rows(lookupRow).Find(lookupValue)
    
    With sheetRange

        'Check if value is in range
        If Not defineRng Is Nothing Then
            cellNumber = defineRng.Column
        Else
            cellNumber = CVErr(xlErrNA)
        End If
        
        'If value is in range then find the column number that the value is in, _
         then convert column number to the column letter
        If Not IsError(cellNumber) Then
            For Each iCell In Intersect(.Columns(cellNumber), .UsedRange)
                Debug.Print iCell.Address
                StringToColumn = Split(Cells(lookupRow, cellNumber).Address, "$")(1)
            Next iCell
        Else
            'Error Message
            msgbox "The string " & lookupValue & " was not found"
        End If
        
    End With
    
End Function

'Returns the number of the last row in a column 
Private Function LastRow(lookupColumn As String, lookupSheet As Worksheet) As long

    LastRow = lookupSheet.Range(lookupColumn & Rows.Count).End(xlUp).Row

End Function

'Returns the number of the last column in a row 
Private Function LastColumn(lookupRow As Integer, lookupSheet As Worksheet) As Integer

    LastColumn = lookupSheet.Cells(lookupRow, Columns.Count).End(xlToLeft).Column
    
End Function

'Delete column with a string in the first row - requires exact match
Sub deleteColumn(stringToDelete As String)

    Dim lastCol As Long
    Dim row As Long
    Dim iCol As Long
    Dim delString As String
    
    delString = stringToDelete
    row = 1
    lastCol = Sheet1.Cells(row, Columns.Count).End(xlToLeft).Column
    
    For iCol = lastCol To 1 Step -1
        If Cells(1, iCol) = delString Then 'You can change this text
            Columns(iCol).Delete
        End If
    Next
End Sub

'Sort up to 10 columns by header provided - requires exact matches
'Additional columns can be sorted by extended this sub
Sub OrderColumns(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10 As String)

    Dim colOrder As Variant
    Dim col As Integer
    Dim search As Range
    Dim index As Integer
        
    colOrder = Array(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
    col = 1
    
    For index = LBound(colOrder) To UBound(colOrder)
        Set search = Rows("1:1").Find(colOrder(index), LookIn:=xlValues, LookAt:=xlWhole, _
            SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
        If Not search Is Nothing Then
            If search.Column <> col Then
                search.EntireColumn.Cut
                Columns(col).Insert Shift:=xlToRight
                Application.CutCopyMode = False
            End If
        col = col + 1
        End If
    Next index
    
End Sub
