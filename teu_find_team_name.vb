Sub find_team_name_teu()
    Dim lastRow As Long
    Dim i As Long
    Dim searchValue As String
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    d.Add "Bernice Lee", "Team AP"
    d.Add "Anita Tsun", "Seal Team"
    d.Add "Eric Hwang", "Team CIF"
    d.Add "Janice Kim", "Team JK"
    d.Add "Julio Garcia", "Team SN"
    d.Add "Philip Asenas", "Team SN"
    d.Add "William Ko", "Team CIF"
    d.Add "Edgar Ayala", "Team JK"
    d.Add "Brian Gong", "Team Mmoss"
    d.Add "Earl Hopper", "Team Seal"
    d.Add "Melody Li", "Team BRL"
    d.Add "Dennis Eng", "Team AP"
    d.Add "Dean Coffman", "Team EW"
    d.Add "Eric Wilson", "Team EW"
    d.Add "Mark Whitlock", "Team SW"
    d.Add "Michael Mendoza", "Seal Team"
    d.Add "Kyle Bloss", "Team SW"
    d.Add "Robert Luce", "Team EW"
    d.Add "Dalia Lara", "Team Tlee"
    d.Add "Nicole Dondi", "Team Tlee"
    d.Add "Mitchell Bennett", "Team Tlee"
    d.Add "Brad Boushey", "Seal Team"
    d.Add "Paul Tiger", "Seal Team"
    d.Add "Jimmy Yim", "Team DK"
    d.Add "Victoria Yow", "Team SN"
    d.Add "Haitham Hamdan", "Team SN"
    d.Add "Steve Shaw", "Team JK"
    d.Add "Joaquin Moreno", "Team JK"
    d.Add "Michelle Moss", "Team Mmoss"
    d.Add "David Kim", "Team DK"
    d.Add "Thomas Lee", "Team Tlee"
    d.Add "Bryan Lee", "Team BRL"
    d.Add "Christy Schmidt", "Team SW"

' determine the last row of data in the range
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in the range
    For i = 1 To lastRow
        ' check if the value in column A is in the dictionary
        searchValue = Cells(i, "E").Value
        If d.Exists(searchValue) Then
            ' if it is, update the value in column C to match the key value from the dictionary
            Cells(i, "B").Value = d(searchValue)
        End If
    Next i
End Sub



