VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuÇalışmaKitabı"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
Workbooks("TEKLİF MG.xlsm").Application.Visible = False
UserForm1.Show 0
End Sub
