VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ANASAYFA"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
UserForm3.Show 0
End Sub
Private Sub CommandButton3_Click()
Sheets("teklif").Copy Before:=Workbooks("TEKL�F MG.xlsm").Sheets(1)
ActiveSheet.Name = "aaa"
TEKL�F.Show 0
End Sub
Private Sub CommandButton4_Click()
m��terilistesi.Show 0
End Sub
Private Sub CommandButton5_Click()
VTEKL�FLER.Show 0
End Sub
Private Sub CommandButton6_Click()
Sheets("fi�").Copy Before:=Workbooks("TEKL�F MG.xlsm").Sheets(1)
ActiveSheet.Name = "bbb"
F��.Show 0
End Sub
Private Sub CommandButton7_Click()
VF��LER.Show 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'Application.Visible = True
Unload UserForm1
ThisWorkbook.Application.Quit
End Sub
