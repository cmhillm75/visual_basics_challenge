Sub stock_change()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim outputRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openValue As Double
    Dim closeValue As Double
    Dim totalChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    Dim processedTickers As Object

    ' Create a dictionary to keep track of processed tickers
    Set processedTickers = CreateObject("Scripting.Dictionary")

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Process sheets for Q1, Q2, Q3, Q4
        
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            
            ' Add titles in row 1 for columns I-L
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Volume"

            ' Find the last row in column A
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

            ' Initialize variables
            ticker = ""
            totalVolume = 0
            openValue = 0
            closeValue = 0
            outputRow = 2

            ' Loop through each row in the worksheet
            For i = 2 To lastRow
                ' Find the ticker symbol change
                If ws.Cells(i, 1).Value <> ticker Then
                    ' If not the first ticker and not already processed, calculate and output the previous ticker's data
                    If ticker <> "" And Not processedTickers.exists(ticker) Then
                        ' Get close value for previous ticker
                        closeValue = ws.Cells(i - 1, 6).Value

                        ' Calculate total change
                        totalChange = closeValue - openValue

                        ' Calculate percentage change
                        If openValue <> 0 Then
                            percentageChange = totalChange / openValue
                        Else
                            percentageChange = 0
                        End If

                        ' Output the results for the previous ticker in columns I - L
                        ws.Cells(outputRow, 9).Value = ticker
                        ws.Cells(outputRow, 10).Value = totalChange
                        ws.Cells(outputRow, 11).Value = Format(percentageChange, "0.00%")
                        ws.Cells(outputRow, 12).Value = totalVolume

                        ' Move to the next row for output
                        outputRow = outputRow + 1
                        ' Mark the ticker as processed
                        processedTickers.Add ticker, True
                    End If

                    ' Reset for the new ticker
                    ticker = ws.Cells(i, 1).Value
                    openValue = ws.Cells(i, 3).Value
                    totalVolume = 0
                End If

                'Add total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            Next i

            ' Output the last ticker's data
            If ticker <> "" And Not processedTickers.exists(ticker) Then
                closeValue = ws.Cells(i - 1, 6).Value
                totalChange = closeValue - openValue
                If openValue <> 0 Then
                    percentageChange = totalChange / openValue
                Else
                    percentageChange = 0
                End If
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = totalChange
                ws.Cells(outputRow, 11).Value = Format(percentageChange, "0.00%")
                ws.Cells(outputRow, 12).Value = totalVolume

                ' Mark the ticker as processed
                processedTickers.Add ticker, True
            End If

            ' Apply conditional formatting for positive and negative changes
            With ws.Range("J2:J" & outputRow - 1)
                .FormatConditions.Delete
                
                ' Add conditional formatting for positive changes
                With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0").Interior
                    .PatternColorIndex = xlAutomatic
                    .ColorIndex = 4  ' Green for positive changes
                End With
                
                ' Add conditional formatting for negative changes
                With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=0").Interior
                    .PatternColorIndex = xlAutomatic
                    .ColorIndex = 3  ' Red for negative changes
                End With
            End With

            ' Clear the dictionary for the next worksheet
            processedTickers.RemoveAll
        End If
    Next ws
End Sub


