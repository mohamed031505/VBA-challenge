Sub Mutliple_Stock_data()


    Dim Summary_row As Integer
    Dim Total_Stock_Volume As Double
    Dim Lastrow As Long
    Dim ws As Worksheet
    Dim Start_price As Double
    Dim Closing_price As Double
    Dim Yearly_change As Double
    Dim Percentage_change As Double
    Dim max_increase As Double
    Dim min_increase As Double
    Dim greatest_total_volume As Double
    Dim greatest_ticker As String
    Dim min_ticker As String
     
     
    
    For Each ws In Worksheets
    
        Summary_row = 2
        Tota_Stock_Volume = 0
        Closing_price = 0
        Yearly_change = 0
        Percentage_change = 0
        Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = " Ticker "
        ws.Cells(1, 10).Value = "Yearly_Change"
        ws.Cells(1, 11).Value = "Percentage_Change"
        ws.Cells(1, 12).Value = "Total_Stock_Volume"
        Start_price = ws.Cells(2, 3).Value
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        max_increase = 0
        max_ticker = ""
        min_increase = 9999999
        greatest_total_volume = 0
        greatest_ticker = ""

        
        
    For i = 2 To Lastrow
    
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ws.Cells(Summary_row, 9).Value = ws.Cells(i, 1).Value
            
            ws.Cells(Summary_row, 12).Value = Total_Stock_Volume
            
            Closing_price = ws.Cells(i, 6).Value
            
            Yearly_change = Closing_price - Start_price
            
            ws.Cells(Summary_row, 10).Value = Yearly_change
            
                    If Yearly_change > 0 Then
        
                         ws.Cells(Summary_row, 10).Interior.ColorIndex = 4
        
                    Else
                        ws.Cells(Summary_row, 10).Interior.ColorIndex = 3
        
                    End If
                        
                        If Start_price > 0 Then
                            
                            Percentage_change = Yearly_change / Start_price
                        Else
                        
                            Percentage_change = 0
                        End If
                        
                        
                        If Percentage_change > max_increase Then
                            max_increase = Percentage_change
                            max_ticker = ws.Cells(i, 1).Value
                  
                    End If

                        If Percentage_change < min_increase Then
                        min_increase = Percentage_change
                        min_ticker = ws.Cells(i, 1).Value
                        
                    End If
                    
                    
                      If Total_Stock_Volume > greatest_total_volume Then
                            greatest_total_volume = Total_Stock_Volume
                            greatest_ticker = ws.Cells(i, 1).Value
                  
                    End If
                    

            ws.Cells(Summary_row, 11).Value = Percentage_change
            
            ws.Cells(Summary_row, 11).NumberFormat = " 0.00% "
            
            
            Start_price = ws.Cells(i + 1, 3).Value
            
            Summary_row = Summary_row + 1
            
            Total_Stock_Volume = 0
            
            
            
        End If
        
        
    Next i
    
    ws.Cells(2, 15).Value = max_ticker
    ws.Cells(2, 16).Value = max_increase
    ws.Cells(2, 16).NumberFormat = "0.00%"
    ws.Cells(3, 15).Value = min_ticker
    ws.Cells(3, 16).Value = min_increase
    ws.Cells(3, 16).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = greatest_total_volume
    ws.Cells(4, 15).Value = greatest_ticker
    Next ws
    

End Sub