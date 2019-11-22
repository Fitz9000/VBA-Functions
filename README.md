# VBA-StringToColumn---2-types


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
