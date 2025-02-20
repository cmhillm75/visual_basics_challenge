'Use the output from change list to create table with greatest changes

Sub summary_table()
    Dim ws As Worksheet
    Dim tempWs As Worksheet
    Dim lastRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim i As Long
    Dim ticker As String
    Dim percentageChange As Double
    Dim totalVolume As Double

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Only process sheets named Q1, Q2, Q3, Q4
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            
            ' Initialize variables for greatest values
            greatestIncrease = -1000
            greatestDecrease = 1000
            greatestVolume = 0

            ' Find the last row in column K and L
            lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

            ' Loop through each row in the worksheet to collect data
            For i = 2 To lastRow
                ticker = ws.Cells(i, 9).Value
                percentageChange = ws.Cells(i, 11).Value
                totalVolume = ws.Cells(i, 12).Value

                ' Check for greatest % increase
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    increaseTicker = ticker
                End If

                ' Check for greatest % decrease
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    decreaseTicker = ticker
                End If

                ' Check for greatest total volume
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTicker = ticker
                End If
            Next i

            ' Output the greatest values to the summary table for each quarter in columsn O - Q
            With ws
                .Cells(1, 15).Value = ""  ' No title for column O
                .Cells(1, 16).Value = "Ticker"
                .Cells(1, 17).Value = "Value"
                .Cells(2, 15).Value = "Greatest % Increase"
                .Cells(3, 15).Value = "Greatest % Decrease"
                .Cells(4, 15).Value = "Greatest Total Volume"
                .Cells(2, 16).Value = increaseTicker
                .Cells(2, 17).Value = Format(Round(greatestIncrease * 100, 2), "0.00") & "%"  ' Greatest % Increase Value, rounded and formatted
                .Cells(3, 16).Value = decreaseTicker
                .Cells(3, 17).Value = Format(Round(greatestDecrease * 100, 2), "0.00") & "%"  ' Greatest % Decrease Value, rounded and formatted
                .Cells(4, 16).Value = volumeTicker
                .Cells(4, 17).Value = greatestVolume
            End With
        End If
    Next ws
'Use the output from change list to create table with greatest changes

Sub summary_table()
    Dim ws As Worksheet
    Dim tempWs As Worksheet
    Dim lastRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim i As Long
    Dim ticker As String
    Dim percentageChange As Double
    Dim totalVolume As Double

    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Only process sheets named Q1, Q2, Q3, Q4
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            
            ' Initialize variables for greatest values
            greatestIncrease = -1000
            greatestDecrease = 1000
            greatestVolume = 0

            ' Find the last row in column K and L
            lastRow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row

            ' Loop through each row in the worksheet to collect data
            For i = 2 To lastRow
                ticker = ws.Cells(i, 9).Value
                percentageChange = ws.Cells(i, 11).Value
                totalVolume = ws.Cells(i, 12).Value

                ' Check for greatest % increase
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    increaseTicker = ticker
                End If

                ' Check for greatest % decrease
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    decreaseTicker = ticker
                End If

                ' Check for greatest total volume
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTicker = ticker
                End If
            Next i

            ' Output the greatest values to the summary table for each quarter in columsn O - Q
            With ws
                .Cells(1, 15).Value = ""  ' No title for column O
                .Cells(1, 16).Value = "Ticker"
                .Cells(1, 17).Value = "Value"
                .Cells(2, 15).Value = "Greatest % Increase"
                .Cells(3, 15).Value = "Greatest % Decrease"
                .Cells(4, 15).Value = "Greatest Total Volume"
                .Cells(2, 16).Value = increaseTicker
                .Cells(2, 17).Value = Format(Round(greatestIncrease * 100, 2), "0.00") & "%"  ' Greatest % Increase Value, rounded and formatted
                .Cells(3, 16).Value = decreaseTicker
                .Cells(3, 17).Value = Format(Round(greatestDecrease * 100, 2), "0.00") & "%"  ' Greatest % Decrease Value, rounded and formatted
                .Cells(4, 16).Value = volumeTicker
                .Cells(4, 17).Value = greatestVolume
            End With
        End If
    Next ws
End Sub