Attribute VB_Name = "Module1"
Sub stockdata()

Dim ticker As String
Dim yearly_change, price_open, price_close, percent_change, volume, lrow, plrow, increase, decrease, vlrow, greatvol As Double

For Each ws In ThisWorkbook.Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    percentage_change = 0
    yearly_change = 0
    volume = 0
    x = 2
    y = 2
    
   
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow

    price_open = ws.Cells(y, 3).Value
            
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            price_close = ws.Cells(i, 6).Value
            yearly_change = (price_close - price_open)
            
            If price_open <> 0 Then
                  percent_change = (Round((yearly_change / price_open) * 100, 2)) & "%"
                  
            Else
                percent_change = 0
            End If
            
            volume = volume + ws.Cells(i, 7).Value
        
            ws.Cells(x, 9).Value = Cells(i, 1).Value
            ws.Cells(x, 10).Value = yearly_change
            ws.Cells(x, 11).Value = percent_change
            ws.Cells(x, 12).Value = volume
            
            x = x + 1
            y = i + 1
            volume = 0
            Else
            volume = volume + ws.Cells(i, 7).Value
        End If
        
    Next i
    
lrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    For j = 2 To lrow
        
        If ws.Cells(j, 10).Value < 0 Then
           ws.Cells(j, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 4
        End If
        
    Next j
 
plrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
increase = 0
decrease = 0

    For k = 2 To plrow
        If increase < ws.Cells(k, 11).Value Then
           increase = ws.Cells(k, 11).Value
           ws.Cells(2, 17).Value = increase
           ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
           ws.Cells(2, 17).Value = ((Round(increase * 100, 2)) & "%")

        ElseIf decrease > ws.Cells(k, 11).Value Then
            decrease = ws.Cells(k, 11).Value
            ws.Cells(3, 17).Value = decrease
            ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(3, 17).Value = ((Round(decrease * 100, 2)) & "%")

            
                        
        End If
    Next k
    
vlrow = ws.Cells(Rows.Count, 12).End(xlUp).Row
greatvol = 0

    For n = 2 To vlrow
        If greatvol < ws.Cells(n, 12).Value Then
            greatvol = ws.Cells(n, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(n, 9).Value
            ws.Cells(4, 17).Value = greatvol
        End If
    Next n
 
    
Next ws
    
        
End Sub


