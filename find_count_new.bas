Attribute VB_Name = "Module7"
Sub find_count_new()
    Dim lastRow As Long
    Dim i As Long
    Dim fso As Object
    Dim oFile As Object
    Dim fileName As String
    
    fileName = "C:\Users\amonroy.lax\Documents\dev\ebay\count_dict.txt" ' replace with the path and name of the file you want to save to'
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    ' determine the last row of data in the range
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in the range
    For i = 1 To lastRow
        ' add the value in column A as a key and the value in column B as the value for that key
        d.Add Cells(i, "A").Value, Cells(i, "B").Value
    Next i
    
    ' create a FileSystemObject to write to the file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(fileName)
    
    ' write the key-value pairs to the file
    For Each k In d.Keys
        oFile.WriteLine k & "," & d(k)
    Next k
    
    ' close the file
    oFile.Close
End Sub
