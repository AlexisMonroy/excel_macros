Sub Check_Rates()
    Dim wbA As Workbook
    Dim wbB As Workbook
    Dim cell As Range
    Dim searchString As String
    Dim searchRange As Range
    Dim foundCell As Range
    
    Set wbA = Workbooks("Bee_International_test.xlsx")
    Set wbB = Workbooks("sbs_test.xlsx")
    Set wsA = wbA.ActiveSheet
    
    

    'Checks for port origin
    For Each cell In wsA.Range("A10:A40")
        If cell.Value = "Ningbo" Then
            wbB.ActiveSheet.Name = "CBP_40HC"
            MsgBox "This is a China port"
            cell.Interior.ColorIndex = 6
        End If
    Next cell

    Set searchRange = wbB.ActiveSheet.Range("A10:B10")
    
    'Checks for port destination and highlights the corresponding cell in workbook B
    For Each searchCell In searchRange
            If InStr(1, searchCell.Value, searchString) > 0 Then
                searchCell.Interior.ColorIndex = 6 'highlight cell in yellow
            foundRow = searchCell.Row 'store the row number of the found cell
            
            'Set the data range to columns B to X of the row where the match was found
            Set dataRange = wbB.ActiveSheet.Range("B" & foundRow & ":X" & foundRow)
            meanValue = Application.WorksheetFunction.Average(dataRange)
            medianValue = Application.WorksheetFunction.Median(dataRange)
            modeValue = Application.WorksheetFunction.Mode(dataRange)
            minValue = Application.WorksheetFunction.Min(dataRange)
            maxValue = Application.WorksheetFunction.Max(dataRange)
            
            'Display a message showing the results
            MsgBox "Mean: " & meanValue & vbCrLf & _
                   "Median: " & medianValue & vbCrLf & _
                   "Mode: " & modeValue & vbCrLf & _
                   "Minimum: " & minValue & vbCrLf & _
                   "Maximum: " & maxValue
        End If
    Next searchCell
End Sub