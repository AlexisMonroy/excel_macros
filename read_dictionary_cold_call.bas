Attribute VB_Name = "Module6"
Sub read_dictionary()
    Dim fso As Object
    Dim oFile As Object
    Dim fileName As String
    Dim line As String
    Dim data() As String
    Dim i As Long
    Dim lastRow As Long
    
    fileName = "C:\Users\amonroy.lax\Documents\dev\ebay\count_dict.txt" ' replace with the path and name of the file you want to read from'
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    ' create a FileSystemObject to read from the file
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.OpenTextFile(fileName)
    
    ' read the key-value pairs from the file and add them to the dictionary
    Do While Not oFile.AtEndOfStream
        line = oFile.ReadLine
        data = Split(line, ",")
        d.Add data(0), data(1)
    Loop
    
    ' close the file
    oFile.Close
    
    ' determine the last row of data in column I
    lastRow = ActiveSheet.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' loop through each row in column I
    For i = 1 To lastRow
        ' check if the value in column I matches a key in the dictionary
        If d.Exists(Cells(i, "I").Value) Then
            ' if it does, add the corresponding value from the dictionary to column D
            Cells(i, "D").Value = d(Cells(i, "I").Value)
        End If
    Next i
    
    ' display a message when the program is finished
    MsgBox "Finished"
End Sub
