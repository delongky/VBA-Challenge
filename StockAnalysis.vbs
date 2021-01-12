Sub StockAnalysis():

'Declare Initial Variables & Set Baseline Variables
Dim totalvolume As Double
Dim RCount As Integer
Dim Start As Integer
Dim yearlychange As Double
Dim percentchange As Double
RCount = 2
Start = 2

'Determine Last Row
RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'Designate Column Headers
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Volume"


'Specifies to ignore zero values for formulas below
For I = 2 To RowCount
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        totalvolume = totalvolume + Cells(I, 7).Value
        If Cells(Start, 3).Value = 0 Then
            For nonzero = Start To I
                If Cells(nonzero, 3).Value <> 0 Then
                    Start = nonzero
                    Exit For
                End If
            Next nonzero
        End If
        
        'Create formulas
        yearlychange = Cells(I, 6) - Cells(Start, 3)
        percentchange = (yearlychange / Cells(Start, 3)) * 100
        Range("I" & RCount).Value = Cells(I, 1).Value
        Range("J" & RCount).Value = yearlychange
        Range("K" & RCount).Value = percentchange
        Range("K" & RCount).NumberFormat = "0.00%"
        Range("L" & RCount).Value = totalvolume
        
        
       'Conditional formatting; cell color change for %change values
        If yearlychange > 0 Then
           Range("J" & RCount).Interior.ColorIndex = 4
        Else
            Range("J" & RCount).Interior.ColorIndex = 3
        End If
        
        RCount = RCount + 1
        totalvolume = 0
        Change = 0
    Else
        totalvolume = totalvolume + Cells(I, 7).Value
    End If
Next I

End Sub
