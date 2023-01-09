Attribute VB_Name = "Module1"
Sub Stock_Test()

For Each ws In Worksheets

Dim ticker As String

Dim total_volume As Double

Dim summary_table_row As Integer
summary_table_row = 2

Dim open_price As Double

Dim close_price As Double


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
    
ws.Columns("I:M").AutoFit

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        total_volume = total_volume + ws.Cells(i, 7).Value
        close_price = ws.Cells(i, 6).Value
        ws.Cells(summary_table_row, 9).Value = ticker
        ws.Cells(summary_table_row, 10).Value = close_price - open_price
        ws.Cells(summary_table_row, 12).Value = total_volume
        
        summary_table_row = summary_table_row + 1
        ticker = ""
        total_volume = 0
        
    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        open_price = ws.Cells(i, 3).Value
    
    Else
        total_volume = total_volume + ws.Cells(i, 7).Value
    End If
    
Next i


Next ws


End Sub
