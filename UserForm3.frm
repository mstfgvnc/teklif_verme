VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10290
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload UserForm3
UserForm2.CommandButton4.Enabled = False
UserForm2.Show
End Sub
Private Sub CommandButton3_Click()
answer = MsgBox(Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 1).Value & " kodlu �r�n� silmek istedi�inize emin misiniz?", vbYesNo + vbQuestion, "�R�N L�STES�")
If answer = vbYes Then
Dim i, a, b As Integer
a = Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 2).Column
b = Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 2).Row
Sheets("�R�NLER").Range("c" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("d" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("e" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("f" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("h" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("�" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("g" & ListBox1.ListIndex + 1).ClearContents
Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 1).ClearContents
If Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 2).Value = "" Then
Else
i = Worksheets("�R�NLER").Range("b655336").End(xlUp).Row
Sheets("�R�NLER").Select
Sheets("�R�NLER").Range(Cells(b, a), Cells(i, a + 7)).Select
Selection.Cut
ActiveSheet.Cells(b - 1, a).Select
ActiveSheet.Paste
End If
Else
End If
End Sub
Private Sub CommandButton4_Click()
On Error Resume Next
UserForm2.TextBox1.Text = Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox2.Text = Sheets("�R�NLER").Range("C" & ListBox1.ListIndex + 1).Value
UserForm2.ComboBox1.Text = Sheets("�R�NLER").Range("D" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox4.Text = Sheets("�R�NLER").Range("E" & ListBox1.ListIndex + 1).Value
UserForm2.ComboBox2.Text = Sheets("�R�NLER").Range("F" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox3.Text = Sheets("�R�NLER").Range("H" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox5.Text = Sheets("�R�NLER").Range("I" & ListBox1.ListIndex + 1).Value
UserForm2.Image1.Picture = LoadPicture(Sheets("�R�NLER").Range("I" & ListBox1.ListIndex + 1).Value)
UserForm2.CommandButton3.Enabled = False
UserForm2.Show
End Sub
Private Sub CommandButton5_Click()
X = InputBox("Miktar giriniz...")
Sheets("aaa").Range("d48").End(xlUp).Offset(1, 0).Value = Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 1).Value
Sheets("aaa").Range("d48").End(xlUp).Offset(0, 1).Value = Sheets("�R�NLER").Range("c" & ListBox1.ListIndex + 1).Value
Sheets("aaa").Range("d48").End(xlUp).Offset(0, 2).Value = Sheets("�R�NLER").Range("h" & ListBox1.ListIndex + 1).Value
Sheets("aaa").Range("d48").End(xlUp).Offset(0, 4).Value = Sheets("�R�NLER").Range("d" & ListBox1.ListIndex + 1).Value
Sheets("aaa").Range("d48").End(xlUp).Offset(0, 5).Value = Sheets("�R�NLER").Range("g" & ListBox1.ListIndex + 1).Value
Sheets("aaa").Range("d48").End(xlUp).Offset(0, 3).Value = X
Unload Me
Dim ts
Set ts = ActiveSheet
a = ActiveSheet.Name
TEKL�F.ListBox1.ColumnCount = 8
TEKL�F.ListBox1.ColumnWidths = "30;80;100;100;50;50;50;50"
TEKL�F.ListBox1.RowSource = a & "!C21:j" & ts.Range("d48").End(xlUp).Row
TEKL�F.ListBox2.RowSource = "aaa!j48"
TEKL�F.ListBox3.RowSource = "aaa!j51"
TEKL�F.ListBox4.RowSource = "aaa!j50"
TEKL�F.ListBox5.RowSource = "aaa!j49"
TEKL�F.ListBox6.RowSource = "aaa!j52"
TEKL�F.Show
End Sub

Private Sub CommandButton6_Click()
X = InputBox("Miktar giriniz...")
Sheets("bbb").Range("C31").End(xlUp).Offset(1, 0).Value = X
Sheets("bbb").Range("C31").End(xlUp).Offset(0, 1).Value = Sheets("�R�NLER").Range("D" & ListBox1.ListIndex + 1).Value
Sheets("bbb").Range("C31").End(xlUp).Offset(0, 2).Value = Sheets("�R�NLER").Range("c" & ListBox1.ListIndex + 1).Value
Unload Me
Dim ts
Set ts = ActiveSheet
a = ActiveSheet.Name
F��.ListBox1.ColumnCount = 4
F��.ListBox1.ColumnWidths = "30;80;100;100"
F��.ListBox1.RowSource = a & "!B5:E" & ts.Range("E31").End(xlUp).Row
F��.Show
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
UserForm2.TextBox1.Text = Sheets("�R�NLER").Range("B" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox2.Text = Sheets("�R�NLER").Range("C" & ListBox1.ListIndex + 1).Value
UserForm2.ComboBox1.Text = Sheets("�R�NLER").Range("D" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox4.Text = Sheets("�R�NLER").Range("E" & ListBox1.ListIndex + 1).Value
UserForm2.ComboBox2.Text = Sheets("�R�NLER").Range("F" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox3.Text = Sheets("�R�NLER").Range("H" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox5.Text = Sheets("�R�NLER").Range("I" & ListBox1.ListIndex + 1).Value
UserForm2.TextBox1.Enabled = False
UserForm2.TextBox2.Enabled = False
UserForm2.TextBox3.Enabled = False
UserForm2.TextBox4.Enabled = False
UserForm2.TextBox5.Enabled = False
UserForm2.ComboBox1.Enabled = False
UserForm2.ComboBox2.Enabled = False
UserForm2.CommandButton1.Enabled = False
UserForm2.CommandButton3.Enabled = False
UserForm2.Image1.Picture = LoadPicture(Sheets("�R�NLER").Range("I" & ListBox1.ListIndex + 1).Value)
UserForm2.CommandButton4.Enabled = False
UserForm2.Show
End Sub
Private Sub UserForm_Initialize()
Dim ts
Set ts = Sheets("�R�NLER")
ListBox1.Clear
ListBox1.ColumnCount = 7
ListBox1.ColumnWidths = "20;80;150;40;40;40;40"
ListBox1.RowSource = "�R�NLER!A1:G" & ts.Range("B" & Rows.Count).End(xlUp).Row
End Sub
