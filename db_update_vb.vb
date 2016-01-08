 Sub test()
 
 Dim header() As String
    Dim i As Integer, j As Integer
    Dim r As Integer, c As Integer
    r = 0
    c = 0
    
    While ActiveCell.Offset(0, c).Value <> ""
      c = c + 1
    Wend
    
    While ActiveCell.Offset(r, 0).Value <> ""
      r = r + 1
    Wend
    
    
    ReDim Preserve header(0 To c)
    
    'Debug.Print "row: " & r & ", column: " & c
    
    For i = 1 To c
        header(i) = Cells(1, i).Value
    Next
    
    Dim v As String, sql As String
    
    Dim result As String
    For i = 2 To r
		result = ""
        sql = "UPDATE " & ActiveSheet.Name & " SET "
        For j = 1 To c
        v = Cells(i, j).Value
        If j = c Then
            sql = sql & header(j) & "='" & v & "' "
        Else
            sql = sql & header(j) & "='" & v & "', "
        End If
            'v = Cells(i, j).Value
            'sql = sql & header(j) & "='" & v & "', "
            'Debug.Print i & j & ": " & Cells(i, j).Interior.ColorIndex
        Next j
        'sql = sql & " where id=" & Cells(i, 1) & ";"
		sql = sql & " where " & Cells(1, 1) & " = " & Cells(i, 1) & ";"
        result = result & sql
        'Debug.Print result
        ActiveSheet.Cells(i, c + 1).Value = result
    Next i

 End Sub

