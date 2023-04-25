Private Type Coordinate
    x As Double
    y As Double
End Type

Sub 计算行列()   ' 字典使用计算行列

  ActiveDocument.Unit = cdrMillimeter
  Set xdict = CreateObject("Scripting.dictionary")
  Set ydict = CreateObject("Scripting.dictionary")
  Dim dot As Coordinate, Offset As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  
  ' 当前选择物件的范围边界
  set_lx = ssr.LeftX: set_rx = ssr.RightX
  set_by = ssr.BottomY: set_ty = ssr.TopY
  ssr(1).GetSize Offset.x, Offset.y
  ' 当前选择物件 ShapeRange 初步排序
  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
'  MsgBox "字典使用计算行列:" & xdict.Count & ydict.Count
  Dim cnt As Long: cnt = 1
  
  ' 遍历字典，输出
  Dim key As Variant
  For Each key In xdict.keys
      dot.x = xdict(key)
      puts dot.x, set_by - Offset.y / 2, cnt
      cnt = cnt + 1
  Next key
  
  cnt = 1
  For Each key In ydict.keys
      dot.y = ydict(key)
      puts set_lx - Offset.x / 2, dot.y, cnt
      cnt = cnt + 1
  Next key
  
End Sub

Private Sub puts(x, y, n)
  Dim st As String
  st = str(n)
  Set s = ActiveLayer.CreateArtisticText(0, 0, st)
  s.CenterX = x: s.CenterY = y
End Sub
