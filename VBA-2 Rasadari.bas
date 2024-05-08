Attribute VB_Name = "Module2"
Sub conditional_formatting():

    For Each ws In Worksheets


    Dim lastrow As Long
    
     ' Determine the last row of data in column A
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
       For i = 2 To lastrow
           If ws.Cells(i, 10) > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
           
           ElseIf ws.Cells(i, 10) < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
           End If
       Next i
    Next ws

End Sub

