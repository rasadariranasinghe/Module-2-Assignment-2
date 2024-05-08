Attribute VB_Name = "Module1"
Sub stock_Data()


    For Each ws In Worksheets
    
    
        Dim TickerName As String
        Dim ClosingPrice As Double
        Dim OpeningPrice As Double
        Dim PC As Double
        Dim QC As Double
        Dim VolumeTotal As Double
        Dim SummaryRow As Integer
        Dim lastrow As Long ' Define lastrow variable
        
        SummaryRow = 2
        VolumeTotal = 0
        
        ' Write headers to the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Determine the last row of data in column A
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastrow
        'To get the opening price of eacg new ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                OpeningPrice = ws.Cells(i, 3).Value
                
            End If
            
           'To get the volume total and closing price of each new ticker
           
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                        
                ClosingPrice = ws.Cells(i, 6).Value
                QC = Round((ClosingPrice - OpeningPrice), 2)
                         
                PC = (QC / OpeningPrice)
                
                
                ws.Range("I" & SummaryRow).Value = TickerName
                ws.Range("L" & SummaryRow).Value = VolumeTotal
                ws.Range("J" & SummaryRow).Value = QC
                ws.Range("J" & SummaryRow).NumberFormat = "0.00"
                ws.Range("K" & SummaryRow).Value = PC
                ws.Range("K" & SummaryRow).NumberFormat = "0.00%"
                
                
                SummaryRow = SummaryRow + 1
                VolumeTotal = 0 ' Reset volume total for the next ticker
            Else
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
               
                
                
            End If
            
        Next i
        
    Next ws
   
End Sub


