VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBA_FORM 
   Caption         =   "Hello_VBA"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "VBA_FORM.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "VBA_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_AQX_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  If Button = 2 Then
    Tools.guideangle ActiveSelectionRange, 0#   ' 右键 0距离贴紧
  ElseIf Shift = fmCtrlMask Then
    Tools.guideangle ActiveSelectionRange, 4    ' 左键安全范围 4mm
  Else
    Tools.guideangle ActiveSelectionRange, -10     ' Ctrl + 鼠标左键
  End If
End Sub

Private Sub CB_BZCC_Click()
  Tools.尺寸标注
End Sub

Private Sub CB_ECWZ_Click()
  Tools.填入居中文字 "你好 CorelVBA!"
End Sub

Private Sub CB_JDZP_Click()
  Tools.角度转平
End Sub

Private Sub CB_PLBZ_Click()
  Tools.批量标注
End Sub

Private Sub CB_PLDYJZ_Click()
  Tools.批量多页居中
End Sub

Private Sub CB_PLWZ_Click()
  Tools.批量居中文字 "CorelVBA批量文字"
End Sub

Private Sub CB_QZJZ_Click()
  Tools.群组居中页面
End Sub

Private Sub CB_SIZESORT_Click()
  Tools.按面积排列 50
End Sub

Private Sub CB_VBA_Click()
  MsgBox "你好 CorelVBA!"
End Sub

Private Sub CB_VBA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
  CB_VBA.BackColor = RGB(255, 0, 0)
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub ZNQZ_Click()
  Tools.智能群组
End Sub
