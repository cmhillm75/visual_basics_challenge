VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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


' Create a reset so we can rerun our original sub each quarter

Sub reset()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    For Each ws In Worksheets
        ' Find the last row and last column with data
        lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Clear contents and formats
        ws.Range(ws.Cells(2, 9), ws.Cells(lastRow, lastCol)).ClearContents
        ws.Range(ws.Cells(2, 9), ws.Cells(lastRow, lastCol)).ClearFormats

        ' Ensure columns O:Q are addressed for the summary table
        ws.Range("O2:Q4").ClearContents
        ws.Range("O2:Q4").ClearFormats
    Next ws
  
  ' Restore the Greatest labels on each worksheet
    For Each ws In Worksheets
        If ws.Name = "Q1" Or ws.Name = "Q2" Or ws.Name = "Q3" Or ws.Name = "Q4" Then
            With ws
                .Cells(2, 15).Value = "Greatest % Increase"
                .Cells(3, 15).Value = "Greatest % Decrease"
                .Cells(4, 15).Value = "Greatest Total Volume"
            End With
        End If
    Next ws
    
End Sub

