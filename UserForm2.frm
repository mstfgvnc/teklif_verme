VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "�R�N EKLE"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13215
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
' RES�M EKLEME MAKROSU
    On Error Resume Next
    Dim vFile As Variant
    vFile = Application.GetOpenFilename("Resim Dosyalar� (*.bmp;*.gif;*.jpg),*.jpg;*.gif;*.gif", 0, "Resim dosyas�", "Open", False)
    If vFile = False Then Exit Sub
    TextResimKisaYol.Text = vFile
    Image1.Picture = LoadPicture(vFile)
    UserForm2.TextBox5.Value = vFile
    Me.Repaint
End Sub
Private Sub CommandButton3_Click()
On Error Resume Next
Dim a, b, C As Double
If UserForm2.TextBox1.Text = "" Then
MsgBox ("L�tfen �r�n Kodunu Giriniz...")
Else
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(1, 0).Value = UserForm2.TextBox1.Text
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 1).Value = UserForm2.TextBox2.Text
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 2).Value = UserForm2.ComboBox1.Text
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 3).Value = UserForm2.TextBox4.Text
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 4).Value = UserForm2.ComboBox2.Text
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 6).Value = UserForm2.TextBox3.Text
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 7).Value = UserForm2.TextBox5.Text
a = CDbl(Worksheets("�R�NLER").Range("Q1").Value)
b = CDbl(Worksheets("�R�NLER").Range("O1").Value)
C = CDbl(UserForm2.TextBox4.Value)
If UserForm2.ComboBox2.Text = "USD" Then
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 5).Value = a * C
Else
End If
If UserForm2.ComboBox2.Text = "EURO" Then
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 5).Value = b * C
Else
End If
If UserForm2.ComboBox2.Text = "TL" Then
Worksheets("�R�NLER").Range("b65536").End(xlUp).Offset(0, 5).Value = C
Else
End If
Unload UserForm2
UserForm3.Show 0
End If
End Sub
Private Sub CommandButton4_Click()
Dim a, b, C As Double
If UserForm2.TextBox1.Text = "" Then
MsgBox ("L�tfen �r�n Kodunu Giriniz...")
Else
Sheets("�R�NLER").Range("B" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.TextBox1.Text
Sheets("�R�NLER").Range("C" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.TextBox2.Text
Sheets("�R�NLER").Range("D" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.ComboBox1.Text
Sheets("�R�NLER").Range("E" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.TextBox4.Text
Sheets("�R�NLER").Range("F" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.ComboBox2.Text
Sheets("�R�NLER").Range("H" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.TextBox3.Text
Sheets("�R�NLER").Range("I" & UserForm3.ListBox1.ListIndex + 1).Value = UserForm2.TextBox5.Text
a = CDbl(Worksheets("�R�NLER").Range("Q1").Value)
b = CDbl(Worksheets("�R�NLER").Range("O1").Value)
C = CDbl(UserForm2.TextBox4.Value)
If UserForm2.ComboBox2.Text = "USD" Then
Sheets("�R�NLER").Range("G" & UserForm3.ListBox1.ListIndex + 1).Value = a * C
Else
End If
If UserForm2.ComboBox2.Text = "EURO" Then
Sheets("�R�NLER").Range("G" & UserForm3.ListBox1.ListIndex + 1).Value = b * C
Else
End If
If UserForm2.ComboBox2.Text = "TL" Then
Sheets("�R�NLER").Range("G" & UserForm3.ListBox1.ListIndex + 1).Value = C
Else
End If
Unload UserForm2
End If
End Sub

Private Sub UserForm_Click()

End Sub
