Attribute VB_Name = "Module1"
' Easy
' Create a script that will loop through each year of stock data
' and grab the total amount of volume each stock had over the year.

' You will also need to display the ticker symbol to coincide with the total volume.

' Your result should look as follows (note: all solution images are for 2015 data).




Sub stock_volume_Easy():

' LOOP THROUGH ALL SHEETS'
    ' --------------------------------------------'
Dim WS As Worksheet
    
For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Determine the Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Add Heading for summary
        Cells(1, "I").Value = "Ticker"
        
        Cells(1, "J").Value = "Total Stock Volume"

    
    
    
    'Set an initial variable for holding the ticker name'
    Dim Ticker_Name As String
    
    'Set an initial variable for holding the total per ticker '
    Dim Ticker_Volume_Total As Double
    
    Ticker_Volume_Total = 0
    'Keep track of the location for each ticker in the summary table'
    
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    'Set i conunter as Long'
    Dim i As Long
    
    
    
    'Loop through all tickers'
    
    For i = 2 To LastRow
        'Check if we are still within the same ticker, if it is  not...'
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            'Set the Ticker Name'
            Ticker_Name = Cells(i, 1).Value
    
            'Add to the Ticker Volume Total'
            Ticker_Volume_Total = Ticker_Volume_Total + Cells(i, 7).Value
    
            'Print the Ticker Name in the Summary Table'
            Range("I" & Summary_Table_Row).Value = Ticker_Name
    
    
            'Print the Ticker Volume to the Summary Table'
            Range("J" & Summary_Table_Row).Value = Ticker_Volume_Total
    
    
            'Add one to the summary table row'
    
            Summary_Table_Row = Summary_Table_Row + 1
    
            'Reset the Ticker Volume Total'
            Ticker_Volume_Total = 0
    
    
            'If the cell immediately following a row is the same ticker...'
    
        Else
    
            'Add to the Ticker Volume Total'
            Ticker_Volume_Total = Ticker_Volume_Total + Cells(i, 7).Value
            
            
       
    
        End If
    
    Next i
    
Next WS


End Sub


