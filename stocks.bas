Attribute VB_Name = "stocks"
Sub StockData():

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

    'label the column headers'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'define variables'
    Dim i As Long
    Dim x As Long 'ends i loop'
    Dim a As Long 'tracks new row'
    Dim opening As Double 'set opening balance'
    Dim closing As Double 'set close balance'
    Dim LastRow As Long
    Dim volume_count As Double
    Dim min_percent As Double
    Dim max_percent As Double
    Dim max_volume As Double
    Dim min_stock As String 'set name of min entity'
    Dim max_stock As String 'set name of max entity'
    Dim max_vol_stock As String 'set name of max volume entity'
        
    'assign variable values before loops'
    a = 2
    i = 2
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    min_percent = 0
    max_percent = 0
    max_vol_stock = 0
    'step 1 - iterate through ticker - sum total volume, % change, etc'
    

    For i = 2 To LastRow
    
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ws.Cells(a, 9).Value = ws.Cells(i, 1).Value
                opening = ws.Cells(i, 3).Value
                volume_count = 0
                volume_count = ws.Cells(i, 7).Value
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                closing = ws.Cells(i, 6).Value
                ws.Cells(a, 10).Value = closing - opening
                    If ws.Cells(a, 10).Value >= 0 Then
                        ws.Cells(a, 10).Interior.ColorIndex = 4
                    Else: ws.Cells(a, 10).Interior.ColorIndex = 3
                    End If
                ws.Cells(a, 11).Value = FormatPercent((closing - opening) / opening)
                    If ws.Cells(a, 11).Value >= 0 Then
                        ws.Cells(a, 11).Interior.ColorIndex = 4
                    Else: ws.Cells(a, 11).Interior.ColorIndex = 3
                    End If
                volume_count = volume_count + ws.Cells(i, 7).Value
                ws.Cells(a, 12).Value = volume_count
                
                'determine summary table'
                'step 2a - min %'
                    If min_percent > ws.Cells(a, 11).Value Then
                         min_percent = ws.Cells(a, 11).Value
                         min_stock = ws.Cells(a, 9).Value
                    Else
                    End If
                
                'step 2b - max %'
                    If max_percent < ws.Cells(a, 11).Value Then
                        max_percent = ws.Cells(a, 11).Value
                        max_stock = ws.Cells(a, 9).Value
                    Else
                    End If
                
                'step 2c - max volume'
                    If max_volume < ws.Cells(a, 12).Value Then
                        max_volume = ws.Cells(a, 12).Value
                        max_vol_stock = ws.Cells(a, 9).Value
                    Else
                    End If
                
                a = a + 1
            ElseIf ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value Then
                volume_count = volume_count + ws.Cells(i, 7).Value
                
            End If
                             
    Next i

    ws.Cells(2, 16).Value = max_stock
    ws.Cells(2, 17).Value = FormatPercent(max_percent)
    
    ws.Cells(3, 16).Value = min_stock
    ws.Cells(3, 17).Value = FormatPercent(min_percent)
    
    ws.Cells(4, 16).Value = max_vol_stock
    ws.Cells(4, 17).Value = max_volume
    
    max_volume = 0
    min_percent = 0
    max_percent = 0
    
Next ws
    
End Sub


