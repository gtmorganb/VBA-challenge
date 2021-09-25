Sub vbachallenge()
'set to do all ws and set variables
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

    Dim ticker As String
    Dim totalvol As Double
    Dim percentchange As Double
    Dim yearlychange As Double
    Dim openval As Double
    Dim closeval As Double
    
    'set summary table row
    Dim summarytablerow As Integer
    summarytablerow = 2
    
    'inital values
    totalvol = 0
    openval = Cells(2, 3).Value
    
    'label headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    
    'find last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop to find values and fill in table
    For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'ticker
    ticker = Cells(i, 1).Value
    Range("I" & summarytablerow).Value = ticker
    
    'yearlychange
    closeval = Cells(i, 6).Value
    yearlychange = closeval - openval
    Range("J" & summarytablerow).Value = yearlychange
    Range("J" & summarytablerow).NumberFormat = "$00.00"
    
    'percent change
    If (openval = 0 And closeval = 0) Then
    percentchange = 0
    ElseIf (openval = 0 And closeval <> 0) Then
    percentchange = 1
    Else
    percentchange = yearlychange / openval
    Range("K" & summarytablerow).Value = percentchange
    Range("K" & summarytablerow).NumberFormat = "0.00%"
    End If
    
    'total volume
    totalvol = totalvol + Cells(i, 7).Value
     Range("L" & summarytablerow).Value = totalvol
    
    'reset values and move row
    totalvol = 0
    openprice = Cells(i + 1, 3)
    summarytablerow = summarytablerow + 1
    
    Else
    totalvol = totalvol + Cells(i, 7).Value
   
    End If
    Next i
      
    'color formatting conditionals
    lasttablerow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For j = 2 To lasttablerow
    If Cells(j, 11) < 0 Then
    Cells(j, 11).Interior.ColorIndex = 3
    Else
    Cells(j, 11).Interior.ColorIndex = 4
    End If
    Next j


    'bonus
    Cells(1, 14).Value = "Greatest % Increase "
    Cells(1, 15).Value = "Greatest % Decrease"
    Cells(1, 16).Value = "Greatest Total Volume"
    
    'new variables
    Dim greatdec As Double
    Dim greatinc As Double
    Dim maxvol As Double
    Dim rng As Range
    
    Set rng = Columns(11)
    greatdec = Application.WorksheetFunction.min(rng)
    Cells(2, 15).Value = greatdec
    Cells(2, 15).NumberFormat = "0.00%"
    
    'MsgBox (min)
    
    greatinc = Application.WorksheetFunction.max(rng)
    Cells(2, 14).Value = greatinc
    Cells(2, 14).NumberFormat = "0.00%"
    
    Dim rng2 As Range
    Set rng2 = Columns(12)
    maxvol = Application.WorksheetFunction.max(rng2)
    Cells(2, 16).Value = maxvol
    
    
Next ws
End Sub
