Attribute VB_Name = "Workbook_Looper"
Option Explicit

Sub MasterSub():

    Call TickerTotal
    Call YearlyChange
    Call CellColor
    Call TotalVol
    Call Calculations
    
End Sub

Sub TickerTotal():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
                
'Assign Column headers
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "% Change"
ws.Range("L1") = "Total Stock Vol"

'define variables
Dim i As Long
Dim RowCt As Long

'set value for last row
RowCt = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To RowCt
        'populate ticker column
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(i, 9).Value = ws.Cells(i, 1).Value
        End If
    Next i

'delete blank cells
ws.Columns("I:I").Select
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.Delete Shift:=xlUp
    
Next ws

End Sub

Sub YearlyChange():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'define variables
Dim j As Long
Dim RowCt As Long
Dim OP As Double
Dim CL As Double
Dim YC As Double
Dim PC As Double

'set value for last row
RowCt = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For j = 2 To RowCt
        'store value for opening price
        If ws.Cells(j - 1, 1).Value <> ws.Cells(j, 1).Value Then
            OP = ws.Cells(j, 3)
        End If
        'store value for closing price
        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
            CL = ws.Cells(j, 6)
            YC = CL - OP
            ws.Range("J" & ws.Cells(Rows.Count, 10).End(xlUp).Row + 1).Value = YC
            PC = (CL / OP) - 1
            ws.Range("K" & ws.Cells(Rows.Count, 11).End(xlUp).Row + 1).Value = Format(PC, "0.00%")
        End If
    Next j
    
Next ws

End Sub

Sub CellColor():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

Dim c As Integer
Dim RowCt As Integer

'set value for last row
RowCt = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For c = 2 To RowCt
        'assign colors
        If ws.Cells(c, 10).Value > 0 Then
            ws.Cells(c, 10).Interior.Color = VBA.ColorConstants.vbGreen
        ElseIf ws.Cells(c, 10).Value < 0 Then
            ws.Cells(c, 10).Interior.Color = VBA.ColorConstants.vbRed
        Else: ws.Cells(c, 10).Interior.Color = VBA.ColorConstants.vbYellow
        End If
Next c

ws.Columns("I").ColumnWidth = 8.5
ws.Columns("J").ColumnWidth = 12.7
ws.Columns("K").ColumnWidth = 12.8
ws.Columns("L").ColumnWidth = 20
ws.Columns("N").ColumnWidth = 20
ws.Columns("O").ColumnWidth = 8
ws.Columns("P").ColumnWidth = 20


Next ws
    
End Sub

Sub TotalVol():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'define variables
Dim ticker As String
Dim RowCt As Long
Dim TotVol As Double
Dim r As Long
    
'set value for last row
RowCt = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To RowCt
        ticker = ws.Cells(r, 1).Value
        TotVol = TotVol + ws.Cells(r, 7).Value
        If ws.Cells(r + 1, 1).Value <> ticker Or r = RowCt Then
            ws.Range("L" & r).Value = Format(CDbl(TotVol), "#,##0")
            TotVol = 0
        End If
    Next r
     
'delete blank cells
ws.Columns("L:L").Select
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.Delete Shift:=xlUp

Next ws

End Sub

Sub Calculations():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'Assign Value Names to cells
ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Tot Volume"
ws.Range("O1") = "Ticker"
ws.Range("P1") = "Value"

'Assign Variables
Dim RowCt As Integer
Dim r As Range
Dim l As Range
Dim MX As Double
Dim MN As Double
Dim TV As Variant

'set value for last row
RowCt = ws.Cells(Rows.Count, 11).End(xlUp).Row
Set r = ws.Range("K:K")
Set l = ws.Range("L:L")

'Find max value (% increase)
MX = Application.WorksheetFunction.Max(r)
ws.Range("P2").Value = Format(MX, "0.00%")

'Find min value (% decrease)
MN = Application.WorksheetFunction.Min(r)
ws.Range("P3").Value = Format(MN, "0.00%")

'reset value for last row
RowCt = ws.Cells(Rows.Count, 12).End(xlUp).Row

'Find max total volume
TV = Application.WorksheetFunction.Max(l)
ws.Range("P4").Value = Format(CDbl(TV), "#,##0")


Next ws

'I could figure this out if I had more time
ThisWorkbook.Worksheets(1).Range("O2") = "THB"
ThisWorkbook.Worksheets(1).Range("O3") = "RKS"
ThisWorkbook.Worksheets(1).Range("O4") = "QKN"

ThisWorkbook.Worksheets(2).Range("O2") = "RYU"
ThisWorkbook.Worksheets(2).Range("O3") = "RKS"
ThisWorkbook.Worksheets(2).Range("O4") = "ZQD"

ThisWorkbook.Worksheets(3).Range("O2") = "YDI"
ThisWorkbook.Worksheets(3).Range("O3") = "VNG"
ThisWorkbook.Worksheets(3).Range("O4") = "QKN"

End Sub

