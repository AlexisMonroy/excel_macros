Sub CheckValues()
    Dim wbA As Workbook, wbB As Workbook
    Dim wsA As Worksheet, wsB As Worksheet
    Dim lastRowA As Long, lastRowB As Long
    Dim i As Long, j As Long
    Dim valueA As Double, valueB As Double
    
    'Set the workbooks and worksheets to use
    Set wbA = Workbooks("WorkbookA.xlsx")
    Set wbB = Workbooks("WorkbookB.xlsx")
    Set wsA = wbA.Worksheets("Sheet1")
    Set wsB = wbB.Worksheets("Sheet1")
    
    'Get the last row in each worksheet
    lastRowA = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, "B").End(xlUp).Row
    
    'Loop through each row in column A of Workbook A
    For i = 1 To lastRowA
        valueA = wsA.Cells(i, "A").Value 'Get the value in column A of Workbook A
        
        'Loop through each row in column B of Workbook B
        For j = 1 To lastRowB
            valueB = wsB.Cells(j, "B").Value 'Get the value in column B of Workbook B
            
            'Check if valueA is less than or equal to valueB
            If valueA <= valueB Then
                'If it is, highlight the cell in Workbook A
                wsA.Cells(i, "A").Interior.Color = vbYellow
                Exit For 'Exit the inner loop
            End If
        Next j 'Go to the next row in Workbook B
    Next i 'Go to the next row in Workbook A
End Sub
