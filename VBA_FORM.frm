VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VBA_FORM 
   Caption         =   "Hello_VBA"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "VBA_FORM.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "VBA_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB_BZCC_Click()
  Tools.�ߴ��ע
End Sub

Private Sub CB_ECWZ_Click()
  Tools.����������� "��� CorelVBA!"
End Sub

Private Sub CB_JDZP_Click()
  Tools.�Ƕ�תƽ
End Sub

Private Sub CB_PLBZ_Click()
  Tools.������ע
End Sub

Private Sub CB_PLWZ_Click()
  Tools.������������ "CorelVBA��������"
End Sub

Private Sub CB_VBA_Click()
  MsgBox "��� CorelVBA!"
End Sub

Private Sub CB_VBA_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  CB_VBA.BackColor = RGB(255, 0, 0)
End Sub

Private Sub ZNQZ_Click()
  Tools.����Ⱥ��
End Sub
