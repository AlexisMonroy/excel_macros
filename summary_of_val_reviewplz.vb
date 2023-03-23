Function getMode(rng As Range) As Variant
    Dim arr As Variant
    Dim dict As Object
    Dim i As Long, n As Long
    Dim modeVal As Variant, modeFreq As Long
    
    arr = Application.Transpose(rng.Value)
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            If dict.Exists(arr(i)) Then
                dict(arr(i)) = dict(arr(i)) + 1
            Else
                dict.Add arr(i), 1
            End If
        End If
    Next i
    
    n = dict.Count
    If n = 0 Then
        getMode = CVErr(xlErrNA)
    ElseIf n = 1 Then
        getMode = dict.Keys()(0)
    Else
        modeFreq = Application.WorksheetFunction.Max(dict.items)
        modeVal = ""
        For i = 0 To n - 1
            If dict.items()(i) = modeFreq Then
                modeVal = modeVal & ", " & dict.Keys()(i)
            End If
        Next i
        getMode = Right(modeVal, Len(modeVal) - 2)
    End If
End Function

.......

'Initialize the variables to hold the summary values
Dim leastAmount As Double, maxAmount As Double, avgAmount As Double, meanAmount As Double, medianAmount As Double, modeAmount As Variant

'Get the last column in row 10 of Workbook B
lastColB = wsB.Cells(10, wsB.Columns.Count).End(xlToLeft).Column

'Summarize the values in row 10 of Workbook B
leastAmount = Application.WorksheetFunction.Min(wsB.Range(wsB.Cells(10, 3), wsB.Cells(10, lastColB)))
maxAmount = Application.WorksheetFunction.Max(wsB.Range(wsB.Cells(10, 3), wsB.Cells(10, lastColB)))
avgAmount = Application.WorksheetFunction.Average(wsB.Range(wsB.Cells(10, 3), wsB.Cells(10, lastColB)))
meanAmount = Application.WorksheetFunction.Median(wsB.Range(wsB.Cells(10, 3), wsB.Cells(10, lastColB)))
modeAmount = getMode(wsB.Range(wsB.Cells(10, 3), wsB.Cells(10, lastColB)))

'Print the summary values to the Immediate Window
Debug.Print "Least amount: " & leastAmount
Debug.Print "Maximum amount: " & maxAmount
Debug.Print "Average amount: " & avgAmount
Debug.Print "Mean amount: " & meanAmount
Debug.Print "Median amount: " & medianAmount
Debug.Print "Mode amount: " & modeAmount
