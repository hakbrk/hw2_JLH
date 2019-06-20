Sub Moderate()

'Dims
Dim LastRow As Double
Dim Stock_Total As Variant
Dim ws As Worksheet
Dim Ticker As String
Dim Summary_Row As Integer
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim First_Row As Double
Dim Last_Price_Row As Double


'Variable initial values
Stock_Total = 0

For Each ws In Worksheets
ws.Activate

    'Set firt summary row
    Summary_Row = 2
    
    'Set the var First_Row to calculate price change
    First_Row = 2
    
    'Add Header
    Range("I1:L1") = Array("Ticker", "Yearly Change", "Percent Change", "Total Volume")
    Range("I1:L1").Font.FontStyle = "Bold"
    
    'Find last row of for loop
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'For Loop
    For i = 2 To LastRow

        If Cells(i, 3) = 0 Then
            First_Row = First_Row + 1
                ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                'Assign Ticker symbol to var Ticker
                Ticker = Cells(i, 1).Value
            
                'Print var Ticker to Summary Row
                Range("I" & Summary_Row).Value = Ticker
            
                'Var Open Price
                Open_Price = Cells(First_Row, 3).Value
            
                'MsgBox (Open_Price)
            
                'Var Close Price
                Close_Price = Cells(i, 6).Value
                'MsgBox (Close_Price)
            
                'Var Yearly Change
                Yearly_Change = (Close_Price - Open_Price)
                If Yearly_Change < 0 Then
                    Range("J" & Summary_Row).Interior.ColorIndex = 3
                    Range("I" & Summary_Row).Interior.ColorIndex = 3
                    Range("K" & Summary_Row).Interior.ColorIndex = 3
                    Range("L" & Summary_Row).Interior.ColorIndex = 3
                    Else
                    Range("J" & Summary_Row).Interior.ColorIndex = 10
                    Range("I" & Summary_Row).Interior.ColorIndex = 10
                    Range("K" & Summary_Row).Interior.ColorIndex = 10
                    Range("L" & Summary_Row).Interior.ColorIndex = 10
                End If
            
                'Print Yearly Change
                Range("J" & Summary_Row).Value = Yearly_Change
                Range("J" & Summary_Row).NumberFormat = "$ 0.00"
            
                'Percent Change
                Percent_Change = Yearly_Change / Open_Price
            
                'Print Percent Change to Summary Table
                Range("K" & Summary_Row).Value = Format(Percent_Change, "Percent")
            
                'Stock volume totals
                Stock_Total = Stock_Total + Cells(i, 7).Value
            
                'Print Stock volume to summary table
                Range("L" & Summary_Row).Value = Format(Stock_Total, "#,###")
            
                'Count Summary Rows
                Summary_Row = Summary_Row + 1
            
                'Reset var Stock Toal to 0
                Stock_Total = 0
            
                'Reset First Row
                First_Row = (i + 1)
            
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
