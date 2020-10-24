Attribute VB_Name = "Module1"
Sub test()

    For Each ws In Worksheets

        'set headers
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly_Change"
        ws.Range("K1").Value = "Percent_Change"
        ws.Range("L1").Value = "Total_Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        'set variable to hold ticker name
        
        Dim ticker As String
        
        'set variable for summary table
        
        Dim table As Integer
        table = 2
        
        'set variable to hold volume
        
        Dim volume As Double
        volume = 0
        
        'set variable for open and close price
        
        Dim open_price As Double
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        
        'set variable to get last row
        
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set opening price for first stock
        
        open_price = ws.Cells(2, 3).Value
        
        'begin loop
        
        For i = 2 To last_row
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value
                
                ws.Range("I" & table).Value = ticker
                
                close_price = ws.Cells(i, 6).Value
                
                yearly_change = close_price - open_price
                
                    If open_price <> 0 Then
                    
                    percent_change = yearly_change / open_price
                    
                    Else: percent_change = 0
                
                    End If
                
                ws.Range("J" & table).Value = yearly_change
                
                    If ws.Range("J" & table).Value > 0 Then
                    
                        ws.Range("J" & table).Interior.ColorIndex = 4
                        
                    Else
                    
                        ws.Range("J" & table).Interior.ColorIndex = 3
                        
                    End If
                
                ws.Range("K" & table).Value = percent_change
                
                ws.Range("K" & table).NumberFormat = "0.00%"
                
                    If ws.Range("K" & table).Value > 0 Then
                    
                        ws.Range("K" & table).Interior.ColorIndex = 4
                        
                    Else
                    
                        ws.Range("K" & table).Interior.ColorIndex = 3
                        
                    End If
                
                open_price = ws.Cells(i + 1, 3).Value
                
                volume = volume + ws.Cells(i, 7).Value
                
                ws.Range("L" & table).Value = volume
                
                table = table + 1
                
                volume = 0
                
            Else
                volume = volume + ws.Cells(i, 7).Value
                    
            End If
            
        Next i
        
        Dim last_row2 As Integer
        
        last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Dim max As Double
        
        max = WorksheetFunction.max(ws.Range("K2:K" & last_row2).Value)
        
        Dim min As Double
        
        min = WorksheetFunction.min(ws.Range("K2:K" & last_row2).Value)
        
        Dim max_volume As Double
        
        max_volume = WorksheetFunction.max(ws.Range("L2:L" & last_row2).Value)
        
        For i = 2 To last_row2
        
            If ws.Cells(i, 11).Value = max Then
            
                ws.Range("P2").Value = max
                
                ws.Range("P2").NumberFormat = "0.00%"
                
                ws.Range("O2").Value = ws.Cells(i, 9).Value
            
            End If
            
            If ws.Cells(i, 11).Value = min Then
            
                ws.Range("P3").Value = min
                
                ws.Range("P3").NumberFormat = "0.00%"
                
                ws.Range("O3").Value = ws.Cells(i, 9).Value
                
            End If
            
            If ws.Cells(i, 12).Value = max_volume Then
            
                ws.Range("P4").Value = max_volume
                
                ws.Range("O4").Value = ws.Cells(i, 9).Value
                
            End If
            
        Next i
        
        ws.Range("A:P").Columns.AutoFit
            
    Next ws
    
End Sub

