Attribute VB_Name = "Module1"
Sub stockData():
   
Dim sheet As Worksheet
    
    For Each sheet In Worksheets
        
    'create titles for the columns we are going to populate data into
    sheet.Cells(1, 9).Value = "Ticker"
    sheet.Cells(1, 10).Value = "Yearly Change"
    sheet.Cells(1, 11).Value = "Percent Change"
    sheet.Cells(1, 12).Value = "Total Stock Volume"
    
    ' variable to hold the ticker name from column A
    ticker = ""
    ' variable to hold the total stock volume
    totalVolume = 0
    ' variable to hold beginning opening price
    openPrice = sheet.Cells(2, 3).Value
    closePrice = 0
    summaryTableRow = 2
    ' use function to find the last row in the sheet
    lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' loop from row 2 in column A out to the last row
        For Row = 2 To lastRow
            
        ' check to see if the ticker changes
            If sheet.Cells(Row + 1, 1).Value <> sheet.Cells(Row, 1).Value Then
                ' set the ticker name
                ticker = sheet.Cells(Row, 1).Value
                ' add the last ticker value from the row
                totalVolume = totalVolume + sheet.Cells(Row, 7).Value
                ' add the ticker to the I column in the summary table row
                sheet.Cells(summaryTableRow, 9).Value = ticker
                
                ' set the closing price for a given ticker
                closePrice = sheet.Cells(Row, 6).Value
                ' print the yearly change in column J
                sheet.Cells(summaryTableRow, 10).Value = (closePrice - openPrice)
                ' print the percent change in column K
                sheet.Cells(summaryTableRow, 11).Value = ((closePrice - openPrice) / openPrice)
                ' format the percent change cells
                sheet.Cells(summaryTableRow, 11).NumberFormat = "#.##%"
                    
                    ' apply conditional formatting to the yearly change column
                    If sheet.Cells(summaryTableRow, 10).Value > 0 Then
                        sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 4
                    ElseIf sheet.Cells(summaryTableRow, 10).Value < 0 Then
                        sheet.Cells(summaryTableRow, 10).Interior.ColorIndex = 3
                    End If
                    
                ' add the total charges to the L column in the summary table row
                sheet.Cells(summaryTableRow, 12).Value = totalVolume
                ' add the currency format to L column
                sheet.Cells(summaryTableRow, 12).Style = "Currency"
                ' go to the next summary table row (add 1 on to the value of the summary table row)
                summaryTableRow = summaryTableRow + 1
                ' reset the ticker total to 0
                totalVolume = 0
                ' reset the new open price for the next ticker
                openPrice = sheet.Cells(Row + 1, 3).Value
                
            Else
                ' if the ticker name stays the same:
                ' add on to the total volume from the G column
                totalVolume = totalVolume + sheet.Cells(Row, 7).Value
            
            End If
            
        Next Row
    
    Next sheet
    
End Sub
