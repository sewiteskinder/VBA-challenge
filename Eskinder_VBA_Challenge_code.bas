Attribute VB_Name = "Module1"
Sub multi_year_stock_data()

    Dim Ticker As String
    
    Dim Total_Stock_Volume As Double
    
    Dim Summary_Table_Row As Integer
        
    Dim Yearly_Open As Double
    
    Dim Yearly_Close As Double
    
    Dim Yearly_Change As Double
        
    Dim Percent_Change As Double
    
    Dim PointerRow As Long
  
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
        PointerRow = 2
        
        ws.Range("i1").Value = "Ticker"
        ws.Range("j1").Value = "Yearly Change"
        ws.Range("k1").Value = "Percent Change"
        ws.Range("l1").Value = "Total Stock Volume"
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("o4").Value = "Greatest Total Volume"
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
        
        RowCount = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        YearlyColor = ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            For j = 2 To RowCount
            
                    If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                    
                        Ticker = ws.Cells(j, 1).Value
                        
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                        
                        ws.Cells(Summary_Table_Row, 9).Value = Ticker
                        
                        Annual_Close = ws.Cells(j, 6).Value
                        
                        Annual_Open = ws.Cells(PointerRow, 3).Value
                        
                        Yearly_Change = Annual_Close - Annual_Open
                        
                        Percent_Change = (Yearly_Change / Annual_Open)
                        
                        ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
       
                        ws.Cells(Summary_Table_Row, 11).Value = "%" & Percent_Change
                        
                        ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                        
                        ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                        
                        Summary_Table_Row = Summary_Table_Row + 1
                        
                        PointerRow = j + 1
                        
                        Total_Stock_Volume = 0
                        
                    Else
                        
                        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
                        
                End If
                
            Next j
            
            
        Dim placeholder_tickermax As String
        
        Dim placeholder_tickermin As String
        
        Dim placeholder_totalvol As String
        
        Dim placeholder_volmax As Double
            placeholder_volmax = 0
        
        Dim placeholder_percentmax As Double
            placeholder_percentmax = 0
        
        Dim placeholder_percentmin As Double
            placeholder_percentmin = 100
        
                For i = 2 To RowCount
                
                    If ws.Cells(i, "K") > placeholder_percentmax Then
                    
                        placeholder_percentmax = ws.Cells(i, "K")
                        
                        placeholder_tickermax = ws.Cells(i, "I")
                        
                    End If
                    
                    If ws.Cells(i, "K") < placeholder_percentmin Then
                    
                        placeholder_percentmin = ws.Cells(i, "K")
                        
                        placeholder_tickermin = ws.Cells(i, "I")
                        
                    End If
                    
                  If ws.Cells(i, "L") > placeholder_volmax Then
                
                    placeholder_volmax = ws.Cells(i, "L")
                    
                    placeholder_totalvol = ws.Cells(i, "I")
                    
                End If
                        
            Next i
                
                ws.Range("P2") = placeholder_tickermax
                
                ws.Range("P3") = placeholder_tickermin
                
                ws.Range("Q2:Q3") = "%" & Percent_Change
                
                ws.Range("Q2") = placeholder_percentmax
                
                ws.Range("Q3") = placeholder_percentmin
                
                ws.Range("P4") = placeholder_totalvol
                
                ws.Range("Q4") = placeholder_volmax
                
            For k = 2 To YearlyColor
            
                If (ws.Cells(k, 10).Value) > 0 Then
                            
                                ws.Cells(k, 10).Interior.ColorIndex = 4
                                
                            ElseIf (ws.Cells(k, 10).Value) < 0 Then
                                
                                ws.Cells(k, 10).Interior.ColorIndex = 3
                                
                            Else
                            
                                ws.Cells(k, 10).Interior.ColorIndex = 2
                                
                End If
            
            Next k
                
        Next ws

End Sub
