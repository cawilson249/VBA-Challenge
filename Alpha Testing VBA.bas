Attribute VB_Name = "Module1"
Sub alpha_Test():

' make worksheet variables
Dim stock_Volume As Long, SummaryTableRow As Integer, TickerStart As Integer, LastClosed As Double, yearlyChange As Long, percentChange As Double



TickerStart = 2
SummaryTableRow = 2

' Loop contol variables
Dim ticker As Integer


' this should help broaden the loop to apply to every sheet
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
    ' create headers for columns
    ws.Range("K1").Value = "Ticker"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent Change"
    ws.Range("N1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("S1").Value = "Greatest % Increase"
    ws.Range("T1").Value = "Greatest % Decrease"
    ws.Range("U1").Value = "Greatest Total Volume"
    ws.Range("V1").Value = "Last Total Volume"

For Row = 2 To lastRow

' if statemnt to loop through the tickers

    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row + 1).Value Then
        
        ticker = ws.Cells(Row, 1).Value
        stock_Volume = stock_Volume + ws.Cells(Row, 7).Value
        ws.Cells(SummaryTableRow, 9).Value = ticker
        ws.Cells(SummaryTableRow, 12).Value = stock_Volume
        SummaryTableRow = SummaryTableRow + 1
    
    Else
        
        stock_Volume = stock_Volume + ws.Cells(Row, 7).Value
    
    End If
    
    yearlyChange = 0
    
    If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
   
        LastClosed = ws.Cells(Row, 6).Value
        
   Else
        
        yearlyChange = ws.Cells(Row + 1, 3).Value - LastClosed
        
        ws.Cells(Row, 10).Value = yearlyChange
        
        percentChange = (yearlyChange / LastClosed)
        
        ws.Cells(SummaryTableRow, 11).Value = percentChange
        SummaryTableRow = SummaryTableRow + 1
    
    End If
    
    ws.Cells(SummaryTableRow, 11).Style = "Percent"
    ws.Cells(SummaryTableRow, 12).Style = "Comma"
    
    
    
    ' Conditional Formatting for the yearly change
    If yearlyChange > 0 Then
        ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(SummaryTableRow, 10).Interior.ColorIndex = 3
        
    End If

Next Row

'Autofit Columns that are being used
    Worksheets("A").Range("A1:Z1").Columns.AutoFit

Next ws

End Sub

