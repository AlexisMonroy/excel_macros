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
    Set wbA = Workbooks("WorkbookA.xlsx")
    Set wbB = Workbooks("WorkbookB.xlsx")
    Set wsA = wbA.Worksheets("Sheet1")
    Set wsB = wbB.Worksheets("Sheet1")
    
    'Get the last row in each worksheet
    lastRowA = wsA.Cells(wsA.Rows.Count, "A").End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, "C").End(xlUp).Row
    
    'Initialize variables
    minAmount = wsB.Cells(10, 3).Value
    maxAmount = wsB.Cells(10, 3).Value
    avgAmount = 0
    meanAmount = 0
    modeAmount = 0
    
    'Copy values from row 10 of Workbook B to an array
    arr = wsB.Range("C10:X10").Value
    
    'Calculate the minimum, maximum, and sum of values
    For k = LBound(arr, 2) To UBound(arr, 2)
        If arr(1, k) < minAmount Then
            minAmount = arr(1, k)
        End If
        If arr(1, k) > maxAmount Then
            maxAmount = arr(1, k)
        End If
        avgAmount = avgAmount + arr(1, k)
    Next k
    
    'Calculate the average and mean of values
    avgAmount = avgAmount / (UBound(arr, 2) - LBound(arr, 2) + 1)
    meanAmount = WorksheetFunction.Average(arr)
    
    'Calculate the median and mode of values
    medianAmount = WorksheetFunction.Median(arr)
    modeAmount = WorksheetFunction.Mode(arr)
    
    'Loop through each row in column A of Workbook A
    For i = 1 To lastRowA
        valueA = wsA.Cells(i, "T").Value 'Get the value in cell T of row i in Workbook A
        
        'Loop through each column in row 10 of Workbook B from column C to X
        For j = 3 To 24
            valueB = wsB.Cells(10, j).Value 'Get the value in row 10, column j in Workbook B
            
            'Check if valueB is greater than valueA
            If valueB > valueA Then
                'Highlight the cell in row i and column j in Workbook B
                wsB.Cells(i, j).Interior.ColorIndex = 6
            End If
        Next j 'Go to the next column in row 10 of Workbook B
    Next i 'Go to the next row in Workbook A
    
    'Display the results in a message box
    message = "Minimum Amount: " & minAmount & vbNewLine
    message = message & "Maximum Amount: " & maxAmount & vbNewLine
    message = message & "Average Amount: " & avgAmount & vbNewLine
