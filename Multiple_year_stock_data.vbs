Sub StockSummary():

    Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
    For Each ws In Worksheets

    
    
        Dim summaryTableRow As Integer
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
    
        
        Dim ticker As String
        Dim startDate As String
        Dim endDate As String
        
        Dim lastRow As Double
        ' Get last row of stock data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Get year start date and end date
        startDate = ws.Cells(2, 2)
        endDate = ws.Cells(lastRow, 2)
        
        Dim yearStartDatePrice As Double
        Dim yearEndDatePrice As Double
        Dim totalVolume As Double
        
        summaryTableRow = 2
        totalVolume = 0
        
        For i = 2 To lastRow
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                ' Add totalVolume, get year end price and save ticker name
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                yearEndDatePrice = ws.Cells(i, 6).Value
                ticker = ws.Cells(i, 1)
                
                ' populate summary table
                ws.Range("I" & summaryTableRow).Value = ticker
                ws.Range("J" & summaryTableRow).Value = yearEndDatePrice - yearStartDatePrice
                ' conditional formating
                If (yearEndDatePrice - yearStartDatePrice < 0) Then
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3
                    
                Else
                    ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4
                End If
                
                ws.Range("K" & summaryTableRow).Value = (yearEndDatePrice - yearStartDatePrice) / yearStartDatePrice
                ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
                ws.Range("L" & summaryTableRow).Value = totalVolume
                
                ' reset data for next ticker
                summaryTableRow = summaryTableRow + 1
                totalVolume = 0
            Else
            
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                If ws.Cells(i, 2) = startDate Then
                    yearStartDatePrice = ws.Cells(i, 3).Value
                End If
                
            End If
        Next i
        
        ' calculate greatest increase and decrease
        
        Dim maxValue As Double
        Dim minValue As Double
        Dim maxTotalValue As Double
        
        maxValue = Application.WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow))
        minValue = Application.WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow))
        maxTotalValue = Application.WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow))
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % increase"
        ws.Range("O3").Value = "Greatest % decrease"
        ws.Range("O4").Value = "Greatest total volume"
        
        For i = 2 To summaryTableRow - 1
        
            If (ws.Cells(i, 11).Value = maxValue) Then
            
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").Value = ws.Cells(i, 11).Value
                ws.Range("Q2").NumberFormat = "0.00%"
                
            ElseIf (ws.Cells(i, 11).Value = minValue) Then
            
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").Value = ws.Cells(i, 11).Value
                ws.Range("Q3").NumberFormat = "0.00%"
                
            ElseIf (ws.Cells(i, 12).Value = maxTotalValue) Then
            
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = ws.Cells(i, 12).Value
                
            End If
                
        Next i
    Next ws

End Sub



