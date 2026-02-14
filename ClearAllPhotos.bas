Sub ClearAllPhotos()
Dim sShape As Shape
Dim i As Long

For i = ActiveSheet.Shapes.Count To 1 Step -1
    Set sShape = ActiveSheet.Shapes(i)
    
    If sShape.Type = msoPicture Then
        sShape.Delete
    End If
Next i

MsgBox "All photos have been cleared!", vbInformation
End Sub

