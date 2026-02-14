Sub InsertMorePhotos()
Dim PicList() As Variant
Dim PicFormat As String
Dim PicRange As Range
Dim sShape As Shape
Dim xRowIndex As Long
Dim xColIndex As Long
Dim lLoop As Long
Dim LastRow As Long

' Find the last picture by checking all shapes
Dim maxRow As Long
maxRow = 2

Dim i As Long
For i = 1 To ActiveSheet.Shapes.Count
    If ActiveSheet.Shapes(i).Type = msoPicture Then
        ' Get the row of the shape based on its top position
        If ActiveSheet.Shapes(i).Top > Cells(maxRow, 2).Top Then
            maxRow = ActiveSheet.Shapes(i).TopLeftCell.Row
        End If
    End If
Next i

' Start inserting photos 3 rows after the last picture
xRowIndex = maxRow + 3
xColIndex = 2

On Error Resume Next
PicList = Application.GetOpenFilename(PicFormat, MultiSelect:=True)

If IsArray(PicList) Then
    For lLoop = LBound(PicList) To UBound(PicList)
        Set PicRange = Cells(xRowIndex, xColIndex)
        Set sShape = ActiveSheet.Shapes.AddPicture2(PicList(lLoop), msoFalse, msoCTrue, PicRange.Left, PicRange.Top, PicRange.Width, PicRange.Height, compress)
        xRowIndex = xRowIndex + 3
    Next
    
    ' Calculate last row
    LastRow = xRowIndex - 3 + 1
    
    ' Update print area
    With ActiveSheet.PageSetup
        .PrintArea = "$A$1:$C$" & LastRow
    End With
    
End If

End Sub
