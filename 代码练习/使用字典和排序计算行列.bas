Private Type Coordinate
    x As Double
    y As Double
End Type

Sub 计算行列()   ' 字典使用计算行列

  ActiveDocument.Unit = cdrMillimeter
  Set xdict = CreateObject("Scripting.dictionary")
  Set ydict = CreateObject("Scripting.dictionary")
  Dim dot As Coordinate
  Dim s As Shape, ssr As ShapeRange
  Set ssr = ActiveSelectionRange
  
  ssr.Sort " @shape1.Top * 100 - @shape1.Left > @shape2.Top * 100 - @shape2.Left"
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
  MsgBox "字典使用计算行列:" & xdict.Count & ydict.Count
  
End Sub

