Sub read_dictionary()
    Dim fso As Object
    Dim oFile As Object
    Dim fileName As String
    Dim line As String
    Dim data() As String
    Dim result As String
    
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
    
    ' build a string with the results of the dictionary
    result = "Results:" & vbCrLf
    For Each k In d.Keys
        result = result & k & ": " & d(k) & vbCrLf
    Next k
    
    ' display the results in a message box
    MsgBox result
End Sub
