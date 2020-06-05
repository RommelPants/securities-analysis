Attribute VB_Name = "Module1"
Sub StockReviewer()
    
    'define variables'
    Dim ticker As String
    Dim counter As Variant
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percentage_changed As Double
    Dim volume As Long
    Dim total As Variant
    
    'Set headers for stock analysis columns'
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "total volume"

    last_row = Cells(Rows.Count, "A").End(xlUp).Row
    
    'reseting rows for loops'
    ticker = ""
    counter = 1
    total = 0
    open_price = 0
    close_price = 0
        
    'peformance anaylysis'
    
    For Row = 2 To last_row
    
        If Cells(Row, 1).Value <> ticker Then
        
            counter = counter + 1
            ticker = Cells(Row, 1).Value
            open_price = Cells(Row, 3).Value
            Cells(counter, 9).Value = ticker
            
            total = Cells(counter, 12) = total
            
        Else
        
            total = total + Cells(Row, 7).Value
            Cells(counter, 12).Value = total
            
        End If
        
        If Cells((Row + 1), 1).Value <> ticker Then
            
                close_price = Cells(Row, 6).Value
            
                yearly_change = close_price - open_price
                    
                Cells(counter, 10).Value = yearly_change
                
                If open_price = 0 Then
                    percent_change = yearly_change
                Else
                    percent_change = yearly_change / open_price
                End If
                                
            End If
                
                Cells(counter, 11).Value = percent_change
                
                'color assignment'
                
                If yearly_change > 0 Then
                Cells(counter, 10).Interior.ColorIndex = 4
                
                ElseIf yearly_change < 0 Then
                Cells(counter, 10).Interior.ColorIndex = 3
                
                End If

    Next Row
    
    'Top winner/loser anaylsis'
    
    Dim top_performer As Variant
    Dim top_loser As Variant
    Dim top_volume As Variant

    Range("O2").Value = "Top Gainer by %"
    Range("O3").Value = "Top Loser by %"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    last_row = Cells(Rows.Count, "I").End(xlUp).Row
    
    top_performer = 0
    top_loser = 0
    top_volume = 0
    
    'Function for reviewing stocks and determining performers on previously noted metrics'
    
    For Row = 2 To last_row:
    
        If Cells(Row, 11).Value > top_performer Then
        
            top_performer = Cells(Row, 11).Value
            
            Range("P2").Value = Cells(Row, 9).Value
            Range("Q2").Value = Cells(Row, 11).Value
            
        End If
    
        If Cells(Row, 11).Value < top_loser Then
        
            top_loser = Cells(Row, 11).Value
            
            Range("P3").Value = Cells(Row, 9).Value
            Range("Q3").Value = Cells(Row, 11).Value
            
        End If
        
        If Cells(Row, 12).Value > top_volume Then
        
            top_volume = Cells(Row, 12).Value
            
            Range("P4").Value = Cells(Row, 9).Value
            Range("Q4").Value = Cells(Row, 12).Value
            
        End If
        
    Next Row
        
End Sub


