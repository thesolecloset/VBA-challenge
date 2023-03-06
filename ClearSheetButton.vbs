Sub ClearSheetButton()

'This For Loop Is Used To Process All The Worksheets In The Workbook
For Each Pgs1 In Worksheets
    Pgs1.Activate
    
    'Will Clear The Contents Of Cells Where Data Will Be Inputted Everytime Button Clicked

    Columns("I").Clear
    Columns("J").Clear
    Columns("K").Clear
    Columns("L").Clear
    Columns("O").Clear
    Columns("P").Clear
    Columns("Q").Clear

Next Pgs1

MsgBox ("I Have Cleared All The Sheets For You")

End Sub


