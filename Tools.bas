Attribute VB_Name = "Tools"
Public Sub 填入居中文字(Str)
  Dim s As Shape
  Set s = ActiveSelection
  x = s.CenterX
  Y = s.CenterY
  
  Set s = ActiveLayer.CreateArtisticText(0, 0, Str)
  s.CenterX = x
  s.CenterY = Y
End Sub

Public Sub 尺寸标注()
  ActiveDocument.Unit = cdrMillimeter
  Set s = ActiveSelection
  x = s.CenterX: Y = s.TopY
  sw = s.SizeWidth: sh = s.SizeHeight
        
  Text = Int(sw) & "x" & Int(sh) & "mm"
  Set s = ActiveLayer.CreateArtisticText(0, 0, Text)
  s.CenterX = x: s.BottomY = Y + 5
End Sub

Public Sub 批量居中文字(Str)
  Dim s As Shape, sr As ShapeRange
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    x = s.CenterX: Y = s.CenterY
    
    Set s = ActiveLayer.CreateArtisticText(0, 0, Str)
    s.CenterX = x: s.CenterY = Y
  Next
End Sub

Public Sub 批量标注()
  ActiveDocument.Unit = cdrMillimeter
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    x = s.CenterX: Y = s.TopY
    sw = s.SizeWidth: sh = s.SizeHeight
          
    Text = Int(sw + 0.5) & "x" & Int(sh + 0.5) & "mm"
    Set s = ActiveLayer.CreateArtisticText(0, 0, Text)
    s.CenterX = x: s.BottomY = Y + 5
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


' 实践应用: 选择物件群组,页面设置物件大小,物件页面居中
Public Function 群组居中页面()
  ActiveDocument.Unit = cdrMillimeter
  Dim OrigSelection As ShapeRange, sh As Shape
  Set OrigSelection = ActiveSelectionRange
  Set sh = OrigSelection.Group
  
  ' MsgBox "选择物件尺寸: " & sh.SizeWidth & "x" & sh.SizeHeight
  ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
  
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

End Function


Public Function 批量多页居中()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True

  ActiveDocument.Unit = cdrMillimeter
  Set sr = ActiveSelectionRange
  total = sr.Count

  '// 建立多页面
  Set doc = ActiveDocument
  doc.AddPages (total - 1)

  Dim sh As Shape
  
  '// 遍历批量物件，放置物件到页面
  For i = 1 To sr.Count
    doc.Pages(i).Activate
    Set sh = sr.Shapes(i)
    ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
 
   '// 物件居中页面
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

  Next i

  ActiveDocument.EndCommandGroup: Application.Optimization = False
  ActiveWindow.Refresh:   Application.Refresh
Exit Function

ErrorHandler:
  Application.Optimization = False
  MsgBox "请先选择一些物件"
  On Error Resume Next
End Function


'// 安全线: 点击一次建立辅助线，再调用清除参考线
Public Function guideangle(actnumber As ShapeRange, cardblood As Integer)
  Dim sr As ShapeRange
  Set sr = ActiveDocument.MasterPage.GuidesLayer.FindShapes(Type:=cdrGuidelineShape)
  If sr.Count <> 0 Then
    sr.Delete
    Exit Function
  End If
  
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter

  With actnumber
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(0, .TopY - cardblood, 0#)
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(0, .BottomY + cardblood, 0#)
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(.LeftX + cardblood, 0, 90#)
    Set s1 = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(.RightX - cardblood, 0, 90#)
  End With
  
End Function



Public Function 按面积排列(space_width As Double)
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrCenter
  
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
  ssr.Sort "@shape1.width * @shape1.height < @shape2.width * @shape2.height"
#Else
' X4 不支持 ShapeRange.sort
#End If

  Dim Str As String, size As String
  For Each sh In ssr
    size = Int(sh.SizeWidth + 0.5) & "x" & Int(sh.SizeHeight + 0.5) & "mm"
    sh.SetSize Int(sh.SizeWidth + 0.5), Int(sh.SizeHeight + 0.5)
    Str = Str & size & vbNewLine
  Next sh

  ActiveDocument.ReferencePoint = cdrTopLeft
  For Each s In ssr
    If cnt > 1 Then s.SetPosition ssr(cnt - 1).LeftX, ssr(cnt - 1).BottomY - space_width
    cnt = cnt + 1
  Next s

'  写文件，可以EXCEL里统计
'  Set fs = CreateObject("Scripting.FileSystemObject")
'  Set f = fs.CreateTextFile("D:\size.txt", True)
'  f.WriteLine str: f.Close

  Str = 分类汇总(Str)
  Debug.Print Str

  Dim s1 As Shape
  Set s1 = ActiveLayer.CreateParagraphText(0, 0, 100, 150, Str, Font:="华文中宋")
End Function
 
'// 实现Excel里分类汇总功能
Private Function 分类汇总(Str As String) As String
  Dim a, b, d, arr
  Str = VBA.Replace(Str, vbNewLine, " ")
  Do While InStr(Str, "  ")
      Str = VBA.Replace(Str, "  ", " ")
  Loop
  arr = Split(Str)

  Set d = CreateObject("Scripting.dictionary")

  For i = 0 To UBound(arr) - 1
    If d.Exists(arr(i)) = True Then
      d.Item(arr(i)) = d.Item(arr(i)) + 1
    Else
       d.Add arr(i), 1
    End If
  Next

  Str = "   规   格" & vbTab & vbTab & vbTab & "数量" & vbNewLine

  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    ' Debug.Print a(i), b(i)
    Str = Str & a(i) & vbTab & vbTab & b(i) & "条" & vbNewLine
  Next

  分类汇总 = Str & "合计总量:" & vbTab & vbTab & vbTab & UBound(arr) & "条" & vbNewLine
End Function



