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
        'sql = "UPDATE " & ActiveSheet.Name & " SET "
		sql = "INSERT INTO " & ActiveSheet.Name & "("
		For k = 1 To c
		If k = c Then
			sql = sql & header(k) & ") VALUES ("
		Else
			sql = sql & header(k) & ","
		End If
		Next k
		
        For j = 1 To c
        v = Cells(i, j).Value
        If j = c Then
            'sql = sql & header(j) & "='" & v & "' "
			sql = sql & "'" & v & "');"
        Else
            'sql = sql & header(j) & "='" & v & "', "
			sql = sql & "'" & v & "',"
        End If
        Next j
        result = result & sql

        ActiveSheet.Cells(i, c + 1).Value = result
    Next i

 End Sub

