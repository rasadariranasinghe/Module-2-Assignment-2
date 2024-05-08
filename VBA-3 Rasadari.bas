Attribute VB_Name = "Module3"
Sub stock_DataPt2():

    For Each ws In Worksheets

        Dim lastrow As Integer ' Define lastrow variable
        Dim MaxNumber As Double
        Dim MinNumber As Double
        Dim MaxVolume As LongLong
        Dim TickerMax As String
        Dim TickerMin As String
        Dim TickerVolume As String
        
     ' Write headers to the summary table
        ws.Range("O2").Value = "Greateset % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
         ' Determine the last row of data in column A
        lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        MaxNumber = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
       
    
        MinNumber = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
        
        
        MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
       
        
        'Write the values in the summary table
        ws.Range("Q2") = MaxNumber
        ws.Range("Q3") = MinNumber
        ws.Range("Q4") = MaxVolume
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        
         'Finding the corresponding ticker name
        For i = 2 To lastrow
            
            If ws.Cells(i, 11).Value = MaxNumber Then
                TickerMax = ws.Cells(i, 9).Value
                
            End If
                
            If ws.Cells(i, 11).Value = MinNumber Then
                TickerMin = ws.Cells(i, 9).Value
                
            End If
                
            If ws.Cells(i, 12).Value = MaxVolume Then
                TickerVolume = ws.Cells(i, 9).Value
                
            End If
          
        Next i
             
            'Write the ticker names in the summary table
                ws.Range("P2") = TickerMax
                ws.Range("P3") = TickerMin
                ws.Range("P4") = TickerVolume
                    
            
    Next ws
                
End Sub





