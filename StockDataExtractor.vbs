Sub StockDataExtractor()

'Declare All Variables Type
Dim LastRow As Long
Dim OpenNum As Double
Dim CloseNum As Double
Dim Change As Double
Dim Percentage As Double
Dim BestPercentage As Double
Dim Vol1 As Double
Dim Vol1Ticker As String
Dim BestVol1 As Double
Dim BestTicker As String
Dim WorstPercentage As Double
Dim WorstTicker As String
Dim Ticker As String
Dim Counter As Integer


'This For Loop Is Used To Process All The Worksheets In The Workbook
For Each Pgs1 In Worksheets
    Pgs1.Activate
    
    'This Will Determine The Last Row In The Data
    LastRow = Pgs1.Cells(Rows.Count, "A").End(xlUp).Row

    'This Will Place The Column Titles On Each WorksheetIn The Workbook For Our Calculated Data
    Pgs1.Cells(1, 9).Value = "Ticker"
    Pgs1.Cells(1, 10).Value = "Yearly Change"
    Pgs1.Cells(1, 11).Value = "Percent Change"
    Pgs1.Cells(1, 12).Value = "Total Stock Volume"
    Pgs1.Cells(2, 15).Value = "Greatest % Increase"
    Pgs1.Cells(3, 15).Value = "Greatest % Decrease"
    Pgs1.Cells(4, 15).Value = "Greatest Total Volume"
    Pgs1.Cells(1, 16).Value = "Ticker"
    Pgs1.Cells(1, 17).Value = "Value"
    
    'Set All Values To 0 For Each Worksheet In The Workbook
    OpenNum = 0
    Change = 0
    Counter = 0
    Ticker = ""
    Vol1 = 0
    Percentage = 0
    
    'For Loop To Find Each Ticker In The Data & Pull Each One Out
    For i = 2 To LastRow

        Ticker = Cells(i, 1).Value
        
        'If Statement To Get The Open Price For Each Ticker We Pull
        If OpenNum = 0 Then
            
            OpenNum = Cells(i, 3).Value
        
        End If
        
        'Calculating The Total Volume For Each Ticker We Pull
        Vol1 = Vol1 + Cells(i, 7).Value
        
        'If Statement To Compare Each Ticker To The First Ticker Pulled & If Not Equal Post Ticker To The To Column For Tickers
        If Cells(i + 1, 1).Value <> Ticker Then
            
            Counter = Counter + 1
            Cells(Counter + 1, 9) = Ticker
            
            'Store The Closing Price For Each Ticker We Pull
            CloseNum = Cells(i, 6)
            
            'Calculate The Total Yearly Change From Opening Price To Closing Price & Post To The Column For Yearly Change
            Change = CloseNum - OpenNum
            Cells(Counter + 1, 10).Value = Change
            
            'Post The Total Volume To The Column For Total Stock Volume
            Cells(Counter + 1, 12).Value = Vol1
            
            'Calculate Percentage & Display The Percentage With Percentage Formatting
            Percentage = (Change / OpenNum)
            Cells(Counter + 1, 11).Value = Format(Percentage, "Percent")
            
            'Set The Opening Number & Volume Back To Zero For Each Ticker In The If Loop So It Will Store The Correct Opening Price
            OpenNum = 0
            Vol1 = 0
            
        End If
        
    Next i
    
    'Set LastRow To Find The Last Row In The New Columns We Created So We Can Find The Best & Worst Percentages & Best Volume
    LastRow = Pgs1.Cells(Rows.Count, "I").End(xlUp).Row
    
    'Set The Variables To The First Data In The List So We Can Compare Them
    BestPercentage = Cells(2, 11).Value
    BestTicker = Cells(2, 9).Value
    WorstPercentage = Cells(2, 11).Value
    WorstTicker = Cells(2, 9).Value
    BestVol1 = Cells(2, 12).Value
    Vol1Ticker = Cells(2, 9).Value
    
    
    'For Loop To Find Compare Each Of The Above Set Variables In The Data & Pull Each One Out & Store Them In The Variables
    For i = 2 To LastRow
    
        'If Statement To Test For Best Percentage
        If Cells(i, 11).Value > BestPercentage Then
            
            BestPercentage = Cells(i, 11).Value
            BestPercentage_Ticker = Cells(i, 9).Value
        
        End If
        
        'If Statement To Test For Worst Percentage
        If Cells(i, 11).Value < WorstPercentage Then
            
            WorstPercentage = Cells(i, 11).Value
            WorstTicker = Cells(i, 9).Value
        
        End If
        
       'If Statement To Test For Best Volume
        If Cells(i, 12).Value > BestVol1 Then
            
            BestVol1 = Cells(i, 12).Value
            Vol1Ticker = Cells(i, 9).Value
        
        End If
        
    Next i
    
    'Display The Variables We Pulled In The Above For Loop
    Pgs1.Cells(2, 16).Value = BestTicker
    Pgs1.Cells(2, 17).Value = Format(BestPercentage, "Percent")
    Pgs1.Cells(3, 16).Value = WorstTicker
    Pgs1.Cells(3, 17).Value = Format(WorstPercentage, "Percent")
    Pgs1.Cells(4, 16).Value = Vol1Ticker
    Pgs1.Cells(4, 17) = BestVol1
    
    'This Will AutoFit All Columns We Used For The Data We Pulled
    Columns(9).AutoFit
    Columns(10).AutoFit
    Columns(11).AutoFit
    Columns(12).AutoFit
    Columns(15).AutoFit
    Columns(16).AutoFit
    Columns(17).AutoFit
    
Next Pgs1

MsgBox ("I Have Extracted All The Data For You On All Sheets")

End Sub
