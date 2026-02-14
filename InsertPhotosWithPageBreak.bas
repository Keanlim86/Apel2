Sub InsertPhotosWithPageBreak()
Dim PicList() As Variant
Dim PicFormat As String
Dim PicRange As Range
Dim sShape As Shape
Dim xRowIndex As Long
Dim xColIndex As Long
Dim lLoop As Long
Dim LastRow As Long
Dim LastUsedRow As Long

ActiveSheet.Cells(2, 2).Select
On Error Resume Next
PicList = Application.GetOpenFilename(PicFormat, MultiSelect:=True)
xColIndex = Application.ActiveCell.Column

If IsArray(PicList) Then
    xRowIndex = Application.ActiveCell.Row
    
    ' Count number of pictures
    PicCount = UBound(PicList) - LBound(PicList) + 1
    
    For lLoop = LBound(PicList) To UBound(PicList)
        Set PicRange = Cells(xRowIndex, xColIndex)
        Set sShape = ActiveSheet.Shapes.AddPicture2(PicList(lLoop), msoFalse, msoCTrue, _
            PicRange.Left, PicRange.Top, PicRange.Width, PicRange.Height, compress)
        xRowIndex = xRowIndex + 3
    Next
    
    ' Calculate last row based on odd/even count
    If PicCount Mod 2 = 1 Then
        ' Odd number of pictures - add extra space
        LastRow = xRowIndex - 3 + 4
    Else
        ' Even number of pictures
        LastRow = xRowIndex - 3 + 1
    End If
    
    ' Insert page break after the last picture
    ActiveSheet.HPageBreaks.Add Before:=Cells(xRowIndex, 1)
    
    ' Update print area
    With ActiveSheet.PageSetup
        .PrintArea = "$A$1:$C$" & LastRow
    End With
    
End If
End Sub
