Sub CheckValues()
    Dim wbA As Workbook, wbB As Workbook
    Dim wsA As Worksheet, wsB As Worksheet
    Dim lastRowA As Long, lastRowB As Long
    Dim i As Long, j As Long
    Dim valueA As Double, valueB As Double
    Dim message As String
    Dim minAmount As Double, maxAmount As Double, avgAmount As Double, meanAmount As Double, medianAmount As Double, modeAmount As Double
    Dim arr() As Variant
    Dim k As Long

    'Set the workbooks and worksheets to use
    Set wbA = Workbooks("Bee_International_test.xlsx")
    Set wbB = Workbooks("sbs_test.xlsx")
    Set wsA = wbA.ActiveSheet
    Set wsB = wbB.ActiveSheet
    



'Get the last row in column B of Workbook B
    lastRowB = wsB.Cells(10, wsB.Columns.Count).End(xlToLeft).Column

    'Get the value in cell T12 of Workbook A
    valueA = wsA.Range("T12").Value

    minAmount = wsB.Cells(10, 3).Value
    maxAmount = wsB.Cells(10, 3).Value
    meanAmount = 0
    
    medianAmount = 0

    arr = wsB.Range("C10:X10").Value

    For k = LBound(arr, 2) To UBound(arr, 2)
        If arr(1, k) < minAmount Then
            minAmount = arr(1, k)
        End If
        If arr(1, k) > maxAmount Then
            maxAmount = arr(1, k)
        End If
    Next k

    meanAmount = WorksheetFunction.Average(arr)
    medianAmount = WorksheetFunction.Median(arr)




'Loop through each column in row 10 of Workbook B
    For j = 1 To lastRowB
        'Check if the value in the cell is a number
        If IsNumeric(wsB.Cells(10, j).Value) Then
            valueB = wsB.Cells(10, j).Value 'Get the value in row 10 of Workbook B, column j
            
            'Check if valueB is greater than valueA
            If valueB > valueA Then
                'Highlight the cell in Workbook B that is greater than valueA
                wsB.Cells(10, j).Interior.Color = vbYellow
                message = "The value in Workbook B, row 10, column " & j & ", is greater than the value in Workbook A, cell T12."
                MsgBox message 'Display message
            Else
                message = "The value in Workbook B, row 10, column " & j & ", is less than or equal to the value in Workbook A, cell T12."
                MsgBox message 'Display message
            End If
        Else
            'Display message if the value in the cell is not a number
            message = "The value in Workbook B, row 10, column " & j & ", is not a number."
            MsgBox message 'Display message
        End If
    Next j 'Go to the next column in row 10 of Workbook B
    
    message = "The minimum amount is " & minAmount & ", the maximum amount is " & maxAmount & ", the mean amount is " & meanAmount & ", the median amount is " & medianAmount & ", and the mode amount is " & modeAmount & "."
End Sub