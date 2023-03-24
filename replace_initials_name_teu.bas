Attribute VB_Name = "Module2"
Sub replace_initials_name_teu()
    Dim lastRow As Long
    Dim i As Long
    Dim searchValue As String
    
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    
    d.Add "blee", "Bernice Lee"
    d.Add "AT", "Anita Tsun"
    d.Add "ehwang", "Eric Hwang"
    d.Add "jakim", "Janice Kim"
    d.Add "jgarcia", "Julio Garcia"
    d.Add "pasenas", "Philip "
    d.Add "wk", "William Ko"
    d.Add "ea", "Edgar Ayala"
    d.Add "BG", "Brian Gong"
    d.Add "EH", "Earl Hopper"
    d.Add "MLI", "Melody Li"
    d.Add "deng", "Dennis Eng"
    d.Add "dcoffman", "Dean Coffman"
    d.Add "EW", "Eric Wilson"
    d.Add "mwhitlock", "Mark Whitlock"
    d.Add "mmendoza", "Michael Mendoza"
    d.Add "kbloss", "Kyle Boss"
    d.Add "rluce", "Robert Luce"
    d.Add "DLARA", "Dalia Lara"
    d.Add "ndondi", "Nicole Dondi"
    d.Add "mbennett", "Mitchell Bennett"
    d.Add "BB", "Brad Boushey"
    d.Add "PTIGER", "Paul Tiger"
    d.Add "jyim", "Jimmy Yim"
    d.Add "vyow", "Victoria Yow"
    d.Add "hhamdan", "Haitham Hamdan"
    d.Add "sshaw", "Steve Shaw"
    d.Add "jmoreno", "Joaquin Moreno"
    d.Add "mmoss", "Michelle Moss"
    d.Add "DKIM", "David Kim"
    d.Add "TLEE", "Thomas Lee"
    d.Add "brlee", "Bryan Lee"
    d.Add "csschmidt", "Bernice Lee"
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' loop through each row in the range
    For i = 1 To lastRow
        ' check if the value in column A is in the dictionary
        searchValue = Cells(i, "E").Value
        If d.Exists(searchValue) Then
            ' if it is, update the value in column C to match the key value from the dictionary
            Cells(i, "E").Value = d(searchValue)
        End If
    Next i
End Sub



