Sub stock()
Dim Final, open_pcn, Change, yearly_change, close_pc, pchange, pchangeI, pchangeD, rvolume, volume, Gvolume As Double
Dim ticker, TickerI, TickerD, TickerV As String
Dim Summary_Table_row As Integer
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate
    Final = Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Table_row = 2
    open_pcn = Cells(2, 3).Value
    ycopen = Cells(2, 3).Value
    pchangeI = 0
    pchangeD = 0
    Gvolume = 0
    For i = 2 To Final
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'Asignar el valor a la variale Ticker y desplegarla en la tambla de resumen
            ticker = Cells(i, 1).Value
            Range("I" & Summary_Table_row).Value = ticker
        'Calcular y desplegar el year change
            ycclose = Cells(i, 6).Value
            yearly_change = ycclose - ycopen
            'Conditional Formating to cells
            If yearly_change >= 0 Then
              Range("J" & Summary_Table_row).Interior.ColorIndex = 4
            Else
                Range("J" & Summary_Table_row).Interior.ColorIndex = 3
            End If
            Range("J" & Summary_Table_row).Value = yearly_change
        'Calcular y desplegar el percent change
            If yearly_change = 0 Then
                pchange = 0
            Else
            pchange = (yearly_change / open_pcn)
            End If
            'Conditional formating to percentage
            Range("K" & Summary_Table_row).NumberFormat = "0%"
            Range("K" & Summary_Table_row).Value = pchange
            'Greatest Increase/Decrease
            If pchange > pchangeI Then
                pchangeI = pchange
                TickerI = ticker
            ElseIf pchange < pchangeD Then
                pchangeD = pchange
                TickerD = ticker
            End If
        'Calcular y desplegar el Total Stock Volume
            rvolume = Cells(i, 7).Value
            volume = volume + rvolume
            Range("L" & Summary_Table_row).Value = volume
            If volume > Gvolume Then
                Gvolume = volume
                TickerV = ticker
            End If
        'Reinicio de las variables acumulativas
            ycopen = Cells(i + 1, 3).Value
            volume = 0
            open_pcn = Cells(i + 1, 3).Value
            Summary_Table_row = Summary_Table_row + 1
            yearly_change = 0
        Else
        'Acumulador Total Stock Volume
            rvolume = Cells(i, 7).Value
            volume = volume + rvolume
            
        End If
            
            
    Next i
    
    Range("Q2").Value = pchangeI
    Range("Q3").Value = pchangeD
    Range("P2").Value = TickerI
    Range("P3").Value = TickerD
    Range("P4").Value = TickerV
    Range("Q4").Value = Gvolume
Next ws
End Sub
