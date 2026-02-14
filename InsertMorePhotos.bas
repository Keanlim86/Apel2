Sub InsertMorePhotos()
Dim PicList() As Variant
Dim PicFormat As String
Dim PicRange As Range
Dim sShape As Shape
Dim xRowIndex As Long
Dim xColIndex As Long
Dim lLoop As Long
Dim LastRow As Long
Dim PicCount As Long
Dim ExistingPicCount As Long
Dim TotalPicCount As Long

' Find the last picture and count existing pictures
Dim maxRow As Long
maxRow = 2
ExistingPicCount = 0
Dim i As Long

For i = 1 To ActiveSheet.Shapes.Count
    If ActiveSheet.Shapes(i).Type = msoPicture Then
        ExistingPicCount = ExistingPicCount + 1  ' <-- Count existing photos
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
    ' Count number of NEW pictures being inserted
    PicCount = UBound(PicList) - LBound(PicList) + 1
    
    ' Calculate TOTAL picture count
    TotalPicCount = ExistingPicCount + PicCount
    
    For lLoop = LBound(PicList) To UBound(PicList)
        Set PicRange = Cells(xRowIndex, xColIndex)
        Set sShape = ActiveSheet.Shapes.AddPicture2(PicList(lLoop), msoFalse, msoCTrue, _
            PicRange.Left, PicRange.Top, PicRange.Width, PicRange.Height, compress)
        xRowIndex = xRowIndex + 3
    Next
    
    ' Calculate last row based on TOTAL odd/even count
    If TotalPicCount Mod 2 = 1 Then
        ' Odd TOTAL number of pictures - add extra space
        LastRow = xRowIndex - 3 + 4
    Else
        ' Even TOTAL number of pictures
        LastRow = xRowIndex - 3 + 1
    End If
    
    ' Update print area
    With ActiveSheet.PageSetup
        .PrintArea = "$A$1:$C$" & LastRow
    End With
    
End If
End Sub
