Sub Easy()

'Dims
Dim LastRow As Double
Dim Stock_Total As Variant
Dim ws As Worksheet
Dim Ticker As String
Dim Summary_Row As Integer

'Variable initial values
Stock_Total = 0

For Each ws In Worksheets
ws.Activate

    'Set firt summary row
    Summary_Row = 2
    
    'Add Header
    Range("I1:J1") = Array("Ticker", "Total Volume")
    Range("I1:J1").Font.FontStyle = "Bold"
    
    'Find last row of for loop
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'For Loop
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Assign Ticker symbol to var Ticker
            Ticker = Cells(i, 1).Value
            
            'Print var Ticker to Summary Row
            Range("I" & Summary_Row).Value = Ticker
            
            'Stock volume totals
            Stock_Total = Stock_Total + Cells(i, 7).Value
            
            'Print Stock volume to summary table
            Range("J" & Summary_Row).Value = Format(Stock_Total, "#,###")
            
            'Count Summary Rows
            Summary_Row = Summary_Row + 1
            
            'Reset var Stock Toal to 0
            Stock_Total = 0
            
        Else
          Stock_Total = Stock_Total + Cells(i, 7).Value
          
        End If
    Next i
    
'Auto fit columns
Cells.Columns.AutoFit

'Center headers
Range("I1:J1").HorizontalAlignment = xlCenter
        
Next ws


End Sub
