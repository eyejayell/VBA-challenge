Attribute VB_Name = "Module1"
Sub Stock_Test()

For Each ws In Worksheets

Dim ticker As String

Dim total_volume As Double

Dim summary_table_row As Integer
summary_table_row = 2

Dim open_price As Double

Dim close_price As Double

Dim yearly_change As Double

Dim percent_change As Double


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
        yearly_change = close_price - open_price
        percent_change = (yearly_change / open_price)
        ws.Cells(summary_table_row, 9).Value = ticker
        ws.Cells(summary_table_row, 10).Value = yearly_change
            If yearly_change > 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summary_table_row, 10).Interior.ColorIndex = 6
            End If
        ws.Cells(summary_table_row, 11).Value = percent_change
        ws.Cells(summary_table_row, 11).NumberFormat = "#.00%"

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

LastRow_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row

greatest_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
greatest_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

For i = 2 To LastRow_Table
    For j = 9 To 12
        
        If greatest_increase = ws.Cells(i, j).Value Then
        ws.Cells(2, 17).Value = ws.Cells(i, j)
        ws.Cells(2, 16).Value = ws.Cells(i, j - 2)
        ws.Cells(2, 17).NumberFormat = "#.00%"
        
        ElseIf greatest_decrease = ws.Cells(i, j).Value Then
        ws.Cells(3, 17).Value = ws.Cells(i, j)
        ws.Cells(3, 16).Value = ws.Cells(i, j - 2)
        ws.Cells(3, 17).NumberFormat = "#.00%"
        
        ElseIf greatest_volume = ws.Cells(i, j).Value Then
        ws.Cells(4, 17).Value = ws.Cells(i, j)
        ws.Cells(4, 16).Value = ws.Cells(i, j - 3)
        
        End If

    Next j
Next i

ws.Columns("O:Q").AutoFit

Next ws


End Sub
