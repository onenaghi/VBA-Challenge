Attribute VB_Name = "Module1"
Sub stock_analysis()

'Label Columns
Range("I1").Value = "Ticker"
Range("J1").Value = "Price Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Define Variable
Dim row As LongLong
Dim lastRow As Long
Dim targetRow As Long
Dim volTotal As LongLong
Dim priceOpen As Variant
'Dim greatPct As Variant

'Define Initial Values
    volTotal = 0
    targetRow = 2
    lastRow = Cells(1, 1).End(xlDown).row
     
'Loop for Data
    For row = 2 To (lastRow - 1)
        
        'Opening Price
        priceOpen = Cells(row, 3).Value
        'Calculate for Total Volume
        volTotal = volTotal + Cells(row, 7).Value
        
        If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
            'Add Tiker Symbol when new ticker value
            Range("I" & targetRow).Value = Cells(row, 1).Value
            'Add yearly change (Closing Price - Opening Price)
            Range("J" & targetRow).Value = Cells(row, 6).Value - priceOpen
            'Add percent change (Closing-Opening/Opening)
            If priceOpen <> 0 Then
            Range("K" & targetRow).Value = FormatPercent(Cells(targetRow, 10) / priceOpen)
            End If
            'Add Total Volume for previous ticker
            Range("L" & targetRow).Value = volTotal
            'Set for next entry
            targetRow = targetRow + 1
            'Reset for next ticker
            volTotal = 0
        End If
        
    Next row

'Loop for Format
For row = 2 To Cells(1, 10).End(xlDown).row
    If Cells(row, 10) > 0 Then
    Cells(row, 10).Interior.ColorIndex = 4
    End If
Next row

For row = 2 To Cells(1, 10).End(xlDown).row
    If Cells(row, 10) < 0 Then
    Cells(row, 10).Interior.ColorIndex = 3
    End If
Next row

'Challenge
'Range("P2").Value = 0

'For row = 2 To Cells(1, 10).End(xlDown).row
    'If Cells(row, 10).Value > greatPct Then
    'greatPct = Cells(row, 10)
    'End If
    
'Next row

End Sub
