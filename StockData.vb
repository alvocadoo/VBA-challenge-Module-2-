Sub StockData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Integer
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        
        ' Initialize summary table headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        ' Initialize summary row
        summaryRow = 2
        
        ' Initialize variables
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        ticker = ws.Cells(2, 1).Value
        openPrice = ws.Cells(2, 3).Value
        totalVolume = 0
        
        ' Process data for each worksheet
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ticker Then
                ' Calculate yearly change and percent change
                closePrice = ws.Cells(i - 1, 6).Value
                yearlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = yearlyChange / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Print summary data
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 12).Value = totalVolume
                
                ' Format percent change
                ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                
                ' Reset variables for next ticker
                summaryRow = summaryRow + 1
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(i, 3).Value
                totalVolume = 0
            End If
            
            ' Accumulate total volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value
        Next i
        
        ' Highlight positive and negative yearly changes
        Dim yearRange As Range
        Set yearRange = ws.Range("J2:J" & summaryRow)
        For Each cell In yearRange
            If cell.Value < 0 Then
                cell.Interior.ColorIndex = 3
            ElseIf cell.Value > 0 Then
                cell.Interior.ColorIndex = 4
            End If
        Next cell
        
        ' Format percent values
        ws.Range("K2:K" & summaryRow).NumberFormat = "0.00%"

'Create a table with the highest and lowest values
Dim PerInc As Double
Dim PerDec As Double
Dim TotVol As Double

'Set ranges to search
Dim searchPercent As Range
Dim searchVolume As Range
Set searchPercent = ws.Range("K2:K" & lastRow)
Set searchVolume = ws.Range("L2:L" & lastRow)

'Find highest and lowest values
PerInc = Application.WorksheetFunction.Max(searchPercent)
PerDec = Application.WorksheetFunction.Min(searchPercent)
TotVol = Application.WorksheetFunction.Max(searchVolume)

'Print highest and lowest values to cells
ws.Cells(2, 16).Value = PerInc
ws.Cells(3, 16).Value = PerDec
ws.Cells(4, 16).Value = TotVol

'Find corresponding tickers for highest and lowest percent changes
Dim TicInc As Double
Dim TicDec As Double
Dim TicTot As Double

TicInc = WorksheetFunction.Match(PerInc, searchPercent, 0)
TicDec = WorksheetFunction.Match(PerDec, searchPercent, 0)
TicTot = WorksheetFunction.Match(TotVol, searchVolume, 0)

'Print corresponding ticker names
ws.Cells(2, 15).Value = ws.Cells(TicInc + 1, 9).Value
ws.Cells(3, 15).Value = ws.Cells(TicDec + 1, 9).Value
ws.Cells(4, 15).Value = ws.Cells(TicTot + 1, 9).Value

'Format percent values
ws.Range("P2:P3").NumberFormat = "0.00%"

    Next ws

End Sub