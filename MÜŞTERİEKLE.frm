VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} M��TER�EKLE 
   Caption         =   "M��TER� SE�-EKLE"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7755
   OleObjectBlob   =   "M��TER�EKLE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "M��TER�EKLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
On Error Resume Next
If M��TER�EKLE.TextBox1.Text = "" Then
MsgBox ("L�tfen Firma Ad�n� Giriniz...")
Else
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(1, 0).Value = M��TER�EKLE.TextBox1.Text
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(0, 1).Value = M��TER�EKLE.TextBox2.Text
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(0, 2).Value = M��TER�EKLE.TextBox7.Text
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(0, 3).Value = M��TER�EKLE.TextBox3.Text
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(0, 4).Value = M��TER�EKLE.TextBox4.Text
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(0, 5).Value = M��TER�EKLE.TextBox5.Text
Worksheets("M��TER�").Range("b65536").End(xlUp).Offset(0, 6).Value = M��TER�EKLE.TextBox6.Text
Dim ds
Set ds = CreateObject("Scripting.FileSystemObject")
ds.CreateFolder ThisWorkbook.Path & "\TEKL�FLER\" & M��TER�EKLE.TextBox1.Value
ds.CreateFolder ThisWorkbook.Path & "\F��LER\" & M��TER�EKLE.TextBox1.Value
Unload M��TER�EKLE
m��terilistesi.Show 0
End If
End Sub
Private Sub CommandButton3_Click()
If M��TER�EKLE.TextBox1.Text = "" Then
MsgBox ("L�tfen �r�n Kodunu Giriniz...")
Else
Sheets("M��TER�").Range("B" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox1.Text
Sheets("M��TER�").Range("C" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox2.Text
Sheets("M��TER�").Range("D" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox7.Text
Sheets("M��TER�").Range("E" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox3.Text
Sheets("M��TER�").Range("F" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox4.Text
Sheets("M��TER�").Range("G" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox5.Text
Sheets("M��TER�").Range("H" & m��terilistesi.ListBox1.ListIndex + 1).Value = M��TER�EKLE.TextBox6.Text
Unload M��TER�EKLE
End If
End Sub

Private Sub UserForm_Click()

End Sub
