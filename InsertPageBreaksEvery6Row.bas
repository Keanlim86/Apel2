Sub InsertPageBreaksEvery6Row()

Dim xLastrow As Long
Dim xWs As Worksheet
Set xWs = Application.ActiveSheet
xRow = 6

xWs.ResetAllPageBreaks
xLastrow = xWs.Range("A1").SpecialCells(xlCellTypeLastCell).Row
For i = xRow + 1 To xLastrow Step xRow
xWs.HPageBreaks.Add Before:=xWs.Cells(i, 1)
Next

With xWs.PageSetup
.Zoom = 115
End With


End Sub
