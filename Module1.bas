Attribute VB_Name = "Module1"
Sub Button1_Click()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Integer
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim Summary_table_row As Integer

For Each ws In alphabetical_testing.xlsx
ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

Summary_table_row = 2

For j = 2 To ws.UsedRange.Rows.Count
If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value And ws.Cells(j, 6) <> 0 Then

year_open = ws.Cells(j, 3).calue
year_close = ws.Cells(j + 261, 6).Value

ticker = ws.Cells(i, 1).Value
vol = ws.Cells(i, 7).Value

yearly_change = year_close - year_open
percent_change = year_close / year_open

ws.Cells(Summary_table_row, 10).Value = yearly_change
ws.Cells(Summary_table_row, 11).Value = percent_change
Summary_table_row = Summary_table_row + 1

year_open = 0
year_close = 0
yearly_change = 0
percent_change = 0

Else

year_open = ws.Cells(j, 3).Value
year_close = ws.Cells(j + 261, 6).Value
yearly_change = year_close - year_open
percent_change = (year_close / year_open)

End If

Next j

ws.Columns("K").NumberFormat = "0.00%"

Dim rg As Range
Dim c As Long
Dim g As Long
Dim color_cell As Range
Set rg = ws.Range("32", Range("j2").End(x1down))
c = rg.Cells.Count
For g = 1 To c
Set color_cell = rg(g)
Select Case color_cell
Case Is < 0
With color_cell
.Interior.Color = vbRed
Case Is > 0
With color_cell
.Interior.Color = vbGreen
End With

End Select
Next i



End Sub


