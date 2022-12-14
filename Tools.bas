Attribute VB_Name = "Tools"
Public Sub �����������(Str)
  Dim s As Shape
  Set s = ActiveSelection
  X = s.CenterX
  Y = s.CenterY
  
  Set s = ActiveLayer.CreateArtisticText(0, 0, Str)
  s.CenterX = X
  s.CenterY = Y
End Sub

Public Sub �ߴ��ע()
  ActiveDocument.Unit = cdrMillimeter
  Set s = ActiveSelection
  X = s.CenterX: Y = s.TopY
  sw = s.SizeWidth: sh = s.SizeHeight
        
  Text = Int(sw) & "x" & Int(sh) & "mm"
  Set s = ActiveLayer.CreateArtisticText(0, 0, Text)
  s.CenterX = X: s.BottomY = Y + 5
End Sub

Public Sub ������������(Str)
  Dim s As Shape, sr As ShapeRange
  Set sr = ActiveSelectionRange
  
  For Each s In sr.Shapes
    X = s.CenterX: Y = s.CenterY
    
    Set s = ActiveLayer.CreateArtisticText(0, 0, Str)
    s.CenterX = X: s.CenterY = Y
  Next
End Sub

Public Sub ������ע()
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

Public Sub ����Ⱥ��()
  Set s1 = ActiveSelectionRange.CustomCommand("Boundary", "CreateBoundary")
  Set brk1 = s1.BreakApartEx

  For Each s In brk1
    Set sh = ActivePage.SelectShapesFromRectangle(s.LeftX, s.TopY, s.RightX, s.BottomY, True)
    sh.Shapes.All.group
    s.Delete
  Next
End Sub


' ʵ��Ӧ��: ѡ�����Ⱥ��,ҳ�����������С,���ҳ�����
Public Function Ⱥ�����ҳ��()
  ActiveDocument.Unit = cdrMillimeter
  Dim OrigSelection As ShapeRange, sh As Shape
  Set OrigSelection = ActiveSelectionRange
  Set sh = OrigSelection.group
  
  ' MsgBox "ѡ������ߴ�: " & sh.SizeWidth & "x" & sh.SizeHeight
  ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
  
#If VBA7 Then
  ActiveDocument.ClearSelection
  sh.AddToSelection
  ActiveSelection.AlignAndDistribute 3, 3, 2, 0, False, 2
#Else
  sh.AlignToPageCenter cdrAlignHCenter + cdrAlignVCenter
#End If

End Function


Public Function ������ҳ����()
  If 0 = ActiveSelectionRange.Count Then Exit Function
  On Error GoTo ErrorHandler
  ActiveDocument.BeginCommandGroup:  Application.Optimization = True

  ActiveDocument.Unit = cdrMillimeter
  Set sr = ActiveSelectionRange
  total = sr.Count

  '// ������ҳ��
  Set doc = ActiveDocument
  doc.AddPages (total - 1)

  Dim sh As Shape
  
  '// ����������������������ҳ��
  For i = 1 To sr.Count
    doc.Pages(i).Activate
    Set sh = sr.Shapes(i)
    ActivePage.SetSize Int(sh.SizeWidth + 0.9), Int(sh.SizeHeight + 0.9)
 
   '// �������ҳ��
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
  MsgBox "����ѡ��һЩ���"
  On Error Resume Next
End Function


'// ��ȫ��: ���һ�ν��������ߣ��ٵ�������ο���
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



Public Function ���������(space_width As Double)
  If 0 = ActiveSelectionRange.Count Then Exit Function
  ActiveDocument.Unit = cdrMillimeter
  ActiveDocument.ReferencePoint = cdrCenter
  
  Set ssr = ActiveSelectionRange
  cnt = 1

#If VBA7 Then
  ssr.Sort "@shape1.width * @shape1.height < @shape2.width * @shape2.height"
#Else
' X4 ��֧�� ShapeRange.sort
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

'  д�ļ�������EXCEL��ͳ��
'  Set fs = CreateObject("Scripting.FileSystemObject")
'  Set f = fs.CreateTextFile("D:\size.txt", True)
'  f.WriteLine str: f.Close

  Str = �������(Str)
  Debug.Print Str

  Dim s1 As Shape
' Set s1 = ActiveLayer.CreateParagraphText(0, 0, 100, 150, Str, Font:="��������")
  X = ssr.FirstShape.LeftX - 100
  Y = ssr.FirstShape.TopY
  Set s1 = ActiveLayer.CreateParagraphText(X, Y, X + 90, Y - 150, Str, Font:="��������")
End Function
 
'// ʵ��Excel�������ܹ���
Private Function �������(Str As String) As String
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

  Str = "   ��   ��" & vbTab & vbTab & vbTab & "����" & vbNewLine

  a = d.keys: b = d.items
  For i = 0 To d.Count - 1
    ' Debug.Print a(i), b(i)
    Str = Str & a(i) & vbTab & vbTab & b(i) & "��" & vbNewLine
  Next

  ������� = Str & "�ϼ�����:" & vbTab & vbTab & vbTab & UBound(arr) & "��" & vbNewLine
End Function


' �����˵������,Ϊ(x1,y1)��(x2,y2) ��ô��Ƕ�a��tanֵ: tana=(y2-y1)/(x2-x1)
' ���Լ���arctan(y2-y1)/(x2-x1), �õ���Ƕ�ֵa
' VB����atn(), ����ֵ�ǻ��ȣ���Ҫ ���� PI /180
Private Function lineangle(x1, y1, x2, y2) As Double
  pi = 4 * VBA.Atn(1) ' ����Բ����
  If x2 = x1 Then
    lineangle = 90: Exit Function
  End If
  lineangle = VBA.Atn((y2 - y1) / (x2 - x1)) / pi * 180
End Function

Public Function �Ƕ�תƽ()
  On Error GoTo ErrorHandler
'  ActiveDocument.ReferencePoint = cdrCenter
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.All

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate -a
    ' sr.LastShape.Delete   '// ɾ���ο���
  End If
ErrorHandler:
End Function

Public Function �Զ���ת�Ƕ�()
  On Error GoTo ErrorHandler
'  ActiveDocument.ReferencePoint = cdrCenter
  Set sr = ActiveSelectionRange
  Set nr = sr.LastShape.DisplayCurve.Nodes.All

  If nr.Count = 2 Then
    x1 = nr.FirstNode.PositionX: y1 = nr.FirstNode.PositionY
    x2 = nr.LastNode.PositionX: y2 = nr.LastNode.PositionY
    a = lineangle(x1, y1, x2, y2): sr.Rotate 90 + a
    sr.LastShape.Delete   '// ɾ���ο���
  End If
ErrorHandler:
End Function


Public Function ��������()
  Set sr = ActiveSelectionRange
  If sr.Count = 2 Then
    X = sr.LastShape.CenterX: Y = sr.LastShape.CenterY
    sr.LastShape.CenterX = sr.FirstShape.CenterX: sr.LastShape.CenterY = sr.FirstShape.CenterY
    sr.FirstShape.CenterX = X: sr.FirstShape.CenterY = Y
  End If
End Function


'//  ===================================================
Private Sub btn_autoalign_byrow_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If get_events("btn_autoalign_byrow", Shift, Button) = "exit" Then Exit Sub
    autogroup("group_lines", 16 + Shift).CreateSelection
End Sub
Private Sub btn_autoalign_bycolumn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If get_events("btn_autoalign_bycolumn", Shift, Button) = "exit" Then Exit Sub
    autogroup("group_lines", 13 + Shift).CreateSelection
End Sub
Private Sub btn_autogroup_byrow_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If get_events("btn_autogroup_byrow", Shift, Button) = "exit" Then Exit Sub
    autogroup("group_lines", 6).CreateSelection
End Sub
Private Sub btn_autogroup_bycolumn_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If get_events("btn_autogroup_bycolumn", Shift, Button) = "exit" Then Exit Sub
    autogroup("group_lines", 3).CreateSelection
End Sub
Private Sub btn_autogroup_bysquare_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If get_events("btn_autogroup_bysquare", Shift, Button) = "exit" Then Exit Sub
    autogroup("group").CreateSelection
End Sub
Private Sub btn_autogroup_byshape_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If get_events("btn_autogroup_byshape", Shift, Button) = "exit" Then Exit Sub
    autogroup("group", 1).CreateSelection
End Sub

Public Sub begin_func(Optional undoname = "nul", Optional units = cdrMillimeter, Optional undogroup = True, Optional optimize = True, Optional sett = "before")
        ActiveDocument.SaveSettings sett
        ActiveDocument.Unit = units
        If undogroup Then ActiveDocument.BeginCommandGroup undoname
        Application.Optimization = optimize
        EventsEnabled = Not optimize
End Sub

Public Sub end_func(Optional undogroup = True, Optional sett = "before")
    cure_app undogroup
    ActiveDocument.RestoreSettings sett
End Sub

Sub cure_app(Optional undogroup = True)
    EventsEnabled = True
    Application.Optimization = False
    Application.Refresh
    DoEvents
    If undogroup Then ActiveDocument.EndCommandGroup
End Sub

Public Function collect_arr(arr, ci, ki)
    lim = UBound(arr)
    For k = 1 To lim
        If arr(ki, k) > 0 Then
            arr(ci, k) = k
            If ki <> ci Then arr(ki, k) = Empty
            If ci <> k And ki <> k Then arr = collect_arr(arr, ci, k)
        End If
    Next k
    'If ki <> ci Then arr(ki, ki) = Empty
    collect_arr = arr
End Function

Public Function autogroup(Optional group As String = "group", Optional shft = 0, Optional sss As Shapes = Nothing, Optional undogroup = True) As ShapeRange
    Dim sr As ShapeRange, sr_all As ShapeRange, os As ShapeRange
    Dim sp As SubPaths
    Dim arr()
    Dim s As Shape
    If sss Is Nothing Then Set os = ActiveSelectionRange Else Set os = sss.All
'On Error GoTo errn
    If ActiveSelection.Shapes.Count > 0 Then
        begin_func "autogroup" & group, cdrMillimeter, undogroup
        gcnt = os.Shapes.Count
        ReDim arr(1 To gcnt, 1 To gcnt)
        Set sr_all = ActiveSelectionRange
        sr_all.RemoveAll
        If group = "group_lines" Then
            For i = 1 To gcnt
                If shft = 3 Or shft = 13 Or shft = 14 Then
                    coord = Int(os.Shapes(i).CenterX)
                Else
                    coord = Int(os.Shapes(i).CenterY)
                End If
                fnd = False
                For k = 1 To gcnt
                    If arr(k, 1) > 0 Then
                        If arr(k, 2) = coord Then
                            arr(k, 1) = arr(k, 1) + 1
                            arr(k, 2 + arr(k, 1)) = i
                            fnd = True
                            Exit For
                        End If
                    Else
                        Exit For
                    End If
                Next k
                If Not fnd Then
                    arr(k, 1) = 1
                    arr(k, 2) = coord
                    arr(k, 3) = i
                End If
            Next i
            Set sr = ActiveSelectionRange
            For i = 1 To gcnt
                If arr(i, 1) > 0 Then
                    sr.RemoveAll
                    For k = 3 To gcnt
                        If arr(i, k) > 0 Then sr.Add os.Shapes(arr(i, k))
                    Next k
                    If sr.Shapes.Count > 0 Then
                        sr.CreateSelection
                        If shft = 13 Then
                            sr.AlignAndDistribute cdrAlignDistributeHNone, cdrAlignDistributeVDistributeSpacing
                        ElseIf shft = 14 Then
                            sr.AlignAndDistribute cdrAlignDistributeHNone, cdrAlignDistributeVDistributeCenter
                        ElseIf shft = 16 Then
                            sr.AlignAndDistribute cdrAlignDistributeHDistributeSpacing, cdrAlignDistributeVNone
                        ElseIf shft = 17 Then
                            sr.AlignAndDistribute cdrAlignDistributeHDistributeCenter, cdrAlignDistributeVNone
                        Else
                            sr.group
                        End If
                        sr_all.AddRange sr
                    End If
                End If
            Next i
        Else
            ReDim arr(1 To gcnt, 1 To gcnt)
            ActiveDocument.Unit = cdrTenthMicron
            sgap = 10
            If shft = 2 Or shft = 3 Or shft = 6 Or shft = 7 Then
                os.RemoveAll
                For Each s In ActiveSelectionRange.Shapes
                    os.Add ActivePage.SelectShapesFromRectangle(s.LeftX - sgap, s.BottomY - sgap, s.RightX + sgap, s.TopY + sgap, True)
                Next s
            End If
            
            For i = 1 To os.Shapes.Count
                Set s1 = os.Shapes(i)
                arr(i, i) = i
                For j = 1 To os.Shapes.Count
                    Set s2 = os.Shapes(j)
                    If s2.LeftX < s1.RightX + sgap And s2.RightX > s1.LeftX - sgap And s2.BottomY < s1.TopY + sgap And s2.TopY > s1.BottomY - sgap Then
                        If shft = 1 Or shft = 3 Or shft = 5 Or shft = 7 Then
                            Set isec = s1.Intersect(s2)
                            If Not isec Is Nothing Then
                                arr(i, j) = j
                                isec.CreateSelection
                                isec.Delete
                            End If
                        Else
                            arr(i, j) = j
                        End If
                    End If
                Next j
            Next i
            
            For i = 1 To gcnt
                arr = collect_arr(arr, i, i)
            Next i
            
            Set sr = ActiveSelectionRange

            For i = 1 To gcnt
                sr.RemoveAll
                inar = 0
                For j = 1 To gcnt
                    If arr(i, j) > 0 Then
                        sr.Add os.Shapes(j)
                        inar = inar + 1
                    End If
                Next j
                If inar > 1 Then
                    If group = "group" Then
                        If shft < 4 Then sr_all.Add sr.group
                    Else
                        If group = "front" Then
                            sr.Sort "@shape1.com.zOrder > @shape2.com.zOrder"
                        ElseIf group = "back" Then
                            sr.Sort "@shape1.com.zOrder < @shape2.com.zOrder"
                        Else
                            sr.Sort "@shape1.width*@shape1.height < @shape2.width*@shape2.height"
                        End If
                        Set fs = sr.FirstShape
                        Set ls = sr.LastShape
                        For Each s In sr.Shapes
                            If Not s Is ls And Not s Is fs Then
                                If group = "autocut" Then
                                    Set isec = ls.Intersect(s)
                                    If Not isec Is Nothing Then
                                        If isec.Curve.Area = s.Curve.Area Then
                                            Set ls = fs.Trim(ls, False)
                                        Else
                                            Set ls = fs.Weld(ls, False)
                                        End If
                                        isec.Delete
                                    End If
                                Else
                                    Set fs = s.Weld(fs, False, False)
                                End If
                            End If
                        Next s
                        If group = "weld" Then
                            Set ls = fs.Weld(ls, False)
                        Else
                            Set ls = fs.Trim(ls, False)
                        End If
                        sr_all.Add ls
                    End If
                Else
                    If sr.Shapes.Count > 0 Then sr_all.AddRange sr
                End If
            Next i
        End If
        Set autogroup = sr_all
    End If
errn:
    end_func undogroup
End Function

Sub auto_cut()
    autogroup("autocut").CreateSelection
End Sub
Sub auto_big_small()
    autogroup("big").CreateSelection
End Sub
Sub auto_group()
    autogroup.CreateSelection
End Sub
Sub auto_weld()
    autogroup("weld").CreateSelection
End Sub
Sub auto_group_lines()
    autogroup("group_lines", 6).CreateSelection
End Sub
Sub auto_group_columns()
    autogroup("group_lines", 3).CreateSelection
End Sub
