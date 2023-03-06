Sub ColorColumn()

'This For Loop Is Used To Process All The Worksheets In The Workbook
For Each Pgs1 In Worksheets
    Pgs1.Activate
    
    LastRow = Pgs1.Cells(Rows.Count, "J").End(xlUp).Row
    
    For i = 2 To LastRow
        
        If Cells(i, 10).Value < 0 Then
        
            Cells(i, 10).Interior.Color = vbRed
            
        End If
        
        If Cells(i, 10).Value >= 0 Then
        
            Cells(i, 10).Interior.Color = vbGreen
            
        End If
        
    Next i

Next Pgs1

MsgBox ("I Have Colored All The Yearly Change Columns For You")
    
End Sub

