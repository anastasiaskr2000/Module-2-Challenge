Sub Stock_market2()

'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for one year
For Each ws In Worksheets

'Define Ticker
Ticker_Row = 1

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
'Dim stock_volume As Double
'stock_volume = 0

'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = ws.Range("C2").Value
Dim close_price As Double
close_price = 0
Dim yearly_change As Double
yearly_change = 0
Dim price_change_percent As Double
price_change_percent = 0
Dim greatest_increase As Double
greatest_increase = 0
Dim greatest_decrease As Double
greatest_decrease = 0
Dim greatest_volume As Double
greatest_volume = 0

'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

    'Ticker symbol output
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Row = Ticker_Row + 1
        Ticker = ws.Cells(i, 1).Value
        ws.Cells(Ticker_Row, "I").Value = Ticker
        
        'Calculate change in Price
        
        close_price = ws.Cells(i, 6).Value
        yearly_change = close_price - open_price
        If open_price = 0 Then
            price_change_percent = 0
        Else
            price_change_percent = (yearly_change / open_price) * 100
        End If
        open_price = ws.Cells(i + 1, 3).Value
        ws.Cells(Ticker_Row, "J").Value = yearly_change
        If yearly_change > 0 Then
        ws.Cells(Ticker_Row, "J").Interior.ColorIndex = 4
        Else
        ws.Cells(Ticker_Row, "J").Interior.ColorIndex = 3
        End If
    
        ws.Cells(Ticker_Row, "K").Value = price_change_percent
        If price_change_percent > 0 Then
        ws.Cells(Ticker_Row, "K").Interior.ColorIndex = 4
        Else
        ws.Cells(Ticker_Row, "K").Interior.ColorIndex = 3
        End If
        
        ws.Cells(Ticker_Row, "L").Value = Ticker_volume
        If yearly_change > greatest_increase Then
            greatest_increase = yearly_change
        End If
        If yearly_change < greatest_decrease Then
            greatest_decrease = yearly_change
        End If
        If Ticker_volume > greatest_volume Then
            greatest_volume = Ticker_volume
        End If
        
        'Reseter
        Ticker_volume = 0
        
        'Calculate the Volume
    Else
        Ticker_volume = Ticker_volume + ws.Cells(i, 7)
        


    End If

Next i

ws.Range("Q2").Value = greatest_increase
ws.Range("Q3").Value = greatest_decrease
ws.Range("Q4").Value = greatest_volume

Next ws

End Sub
