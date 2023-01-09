Attribute VB_Name = "Module1"
Sub Stock_Test()

Dim ticker As String

Dim total_value As Double

Dim summary_table_row As Integer
summary_table_row = 2

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
    
ws.Columns("I:M").AutoFit

For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        Range("I" & summary_table_row).Value = ticker
        summary_table_row = summary_table_row + 1
        ticker = ""
        
    End If
    
Next i

Next ws


End Sub
