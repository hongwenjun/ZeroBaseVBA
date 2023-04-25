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
  
  For Each s In ssr
    dot.x = s.CenterX: dot.y = s.CenterY
    If xdict.Exists(Int(dot.x)) = False Then xdict.Add Int(dot.x), dot.x
    If ydict.Exists(Int(dot.y)) = False Then ydict.Add Int(dot.y), dot.y
  Next s
  
  Dim keys() As Variant
  keys = xdict.keys
  ' 使用 Sort 函数对数组进行排序
  ArraySort keys
  ' 遍历排序后的键，并按照键的顺序访问字典中的元素
  Dim key As Variant
  For Each key In keys
      Debug.Print key, xdict(key)
  Next key

  Debug.Print "字典使用计算行列:" & xdict.Count, ydict.Count
  
End Sub

'// 对数组进行排序[单维]
Public Function ArraySort(src As Variant) As Variant
  Dim out As Long, i As Long, tmp As Variant
  For out = LBound(src) To UBound(src) - 1
    For i = out + 1 To UBound(src)
      If src(out) > src(i) Then
        tmp = src(i): src(i) = src(out): src(out) = tmp
      End If
    Next i
  Next out
  
  ArraySort = src
End Function
