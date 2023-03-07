Attribute VB_Name = "Module1"
Sub StockAnalysis()
    'Define variables
    Dim TickerSymbol As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    'Set initial values
    OpeningPrice = 0
    ClosingPrice = 0
    YearlyChange = 0
    PercentChange = 0
    TotalVolume = 0
    'Loop through all worksheets in the workbook
    For Each ws In Worksheets
        'Set column headers
        ws.Range("I1").Value = "Ticker Symbol"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        'Find last row in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Loop through all rows in the worksheet
        For i = 2 To LastRow
            'Check if TickerSymbol has changed
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'Set TickerSymbol
                TickerSymbol = ws.Cells(i, 1).Value
                'Set OpeningPrice
                OpeningPrice = ws.Cells(i, 3).Value
            End If
            'Check if TickerSymbol is the same
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                'Add TotalVolume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            Else
                'Set ClosingPrice
                ClosingPrice = ws.Cells(i, 6).Value
                'Calculate YearlyChange
                YearlyChange = ClosingPrice - OpeningPrice
                'Check for divide by zero error
                If OpeningPrice = 0 Then
                    PercentChange = 0
                Else
                    'Calculate PercentChange
                    PercentChange = (YearlyChange / OpeningPrice) * 100
                End If
                'Add TotalVolume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                'Output results
                j = 2
                ws.Range("I" & j).Value = TickerSymbol
                ws.Range("J" & j).Value = YearlyChange
                ws.Range("K" & j).Value = PercentChange
                ws.Range("L" & j).Value = TotalVolume
                'Reset values for next iteration
                TickerSymbol = ""
                OpeningPrice = 0
                ClosingPrice = 0
                YearlyChange = 0
                PercentChange = 0
                TotalVolume = 0
                'Move to next row
                j = j + 1
            End If
        Next i
        'Autofit columns
        ws.Columns("I:L").AutoFit
    Next ws
End Sub
