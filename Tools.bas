Attribute VB_Name = "Tools"
Public Sub 填入居中文字(str)
  Dim s As Shape
  Set s = ActiveSelection
  X = s.CenterX
  Y = s.CenterY
  
  Set s = ActiveLayer.CreateArtisticText(0, 0, str)
  s.CenterX = X
  s.CenterY = Y
End Sub

Public Sub 尺寸标注()
  ActiveDocument.Unit = cdrMillimeter
  Set s = ActiveSelection
  X = s.CenterX: Y = s.TopY
  sw = s.SizeWidth: sh = s.SizeHeight
        
  Text = Int(sw) & "x" & Int(sh) & "mm"
  Set s = ActiveLayer.CreateArtisticText(0, 0, Text)
  s.CenterX = X: s.BottomY = Y + 5
End Sub

Public Sub 批量居中文字(str)
  Dim s As Shape, sr As ShapeRange
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    X = s.CenterX: Y = s.CenterY
    
    Set s = ActiveLayer.CreateArtisticText(0, 0, str)
    s.CenterX = X: s.CenterY = Y
  Next
End Sub

Public Sub 批量标注()
  ActiveDocument.Unit = cdrMillimeter
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    X = s.CenterX: Y = s.TopY
    sw = s.SizeWidth: sh = s.SizeHeight
          
    Text = Int(sw + 0.5) & "x" & Int(sh + 0.5) & "mm"
    Set s = ActiveLayer.CreateArtisticText(0, 0, Text)
    s.CenterX = X: s.BottomY = Y + 5
  Next
End Sub

Public Sub 智能群组()
  Set s1 = ActiveSelectionRange.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx

  For Each s In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, True)
    sh.Shapes.All.Group
    s.Delete
  Next
End Sub

Private Function 对角线角度(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
  pi = 4 * VBA.Atn(1) ' 计算圆周率'
  对角线角度 = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
End Function

Public Sub 角度转平()
  ActiveDocument.ReferencePoint = cdrCenter
  Dim sr As ShapeRange '定义物件范围
  Set sr = ActiveSelectionRange

  Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
  Dim Shift As Long
  Dim b As Boolean

  b = ActiveDocument.GetUserArea(x1, y1, x2, y2, Shift, 10, False, 306)
  If Not b Then
    a = 对角线角度(x1, y1, x2, y2)
    sr.Rotate -a
  End If
End Sub
