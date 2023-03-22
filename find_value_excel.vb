Sub TESTTESTTEST()
    Dim lastRow As Long
    Dim i As Long
    Dim searchValue As String
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    d.Add "Taiwan", "HERE"
  

' determine the last row of data in the range
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in the range
    For i = 1 To lastRow
        ' check if the value in column A is in the dictionary
        searchValue = Cells(i, "A").Value
        If d.Exists(searchValue) Then
            ' if it is, update the value in column C to match the key value from the dictionary
            Cells(i, "C").Value = d(searchValue)
        End If
    Next i
End Sub



