
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