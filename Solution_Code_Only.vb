Sub TickerVolumeSummaryTable():
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        Dim lastrow As Double
        Dim i As Double
        Dim Ticker_Name As String
        Dim Ticker_Volume As Double
        Dim Sum_Table_Row As Long
        Dim initial As Double
        Dim final As Double
        Dim perChange As Double
        Dim yearChange As Double
        Dim j As Integer
                
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Ticker_Volume = 0
        Ticker_Name = ""
        Sum_Table_Row = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
        initial = ws.Cells(2, 3).Value
        yearChange = 0
        
        For i = 2 To lastrow
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                Ticker_Name = ws.Cells(i, 1).Value
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & Sum_Table_Row).Value = Ticker_Name
                ws.Range("L" & Sum_Table_Row).Value = Ticker_Volume
                
                final = ws.Cells(i, 6).Value
                yearChange = (final - initial)
                
                If initial <> 0 Then
                    perChange = yearChange / initial
                Else
                    perChange = 0
                End If
                
                ws.Range("J" & Sum_Table_Row).Value = yearChange
                ws.Range("K" & Sum_Table_Row).Value = perChange
                
                Ticker_Volume = 0
                Sum_Table_Row = Sum_Table_Row + 1
                initial = ws.Cells(i + 1, 3).Value
                final = 0
                
            Else
                Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
                            
            End If
        
        Next i
           
        For j = 2 To Sum_Table_Row - 1
            
            If ws.Range("J" & j).Value < 0 Then
                ws.Range("J" & j).Interior.ColorIndex = 3
            
            Else
                ws.Range("J" & j).Interior.ColorIndex = 5
                
            End If
                            
        Next j
           
           
           
           
    Next ws
    
End Sub

