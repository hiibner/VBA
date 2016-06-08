Attribute VB_Name = "ExtractHL-Image.bas"
Sub ConvertHLShapes()
Dim shp As Shape
Dim sTemp As String

For Each shp In ActiveSheet.Shapes
  sTemp = ""
  On Error Resume Next 'go to next shape if no hyperlink
    sTemp = shp.Hyperlink.Address
  On Error GoTo 0
    If sTemp <> "" Then
      shp.TopLeftCell.Value = sTemp
      shp.Delete
    End If
Next
End Sub
