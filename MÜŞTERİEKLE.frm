VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MÜÞTERÝEKLE 
   Caption         =   "MÜÞTERÝ SEÇ-EKLE"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7755
   OleObjectBlob   =   "MÜÞTERÝEKLE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MÜÞTERÝEKLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
On Error Resume Next
If MÜÞTERÝEKLE.TextBox1.Text = "" Then
MsgBox ("Lütfen Firma Adýný Giriniz...")
Else
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(1, 0).Value = MÜÞTERÝEKLE.TextBox1.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 1).Value = MÜÞTERÝEKLE.TextBox2.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 2).Value = MÜÞTERÝEKLE.TextBox7.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 3).Value = MÜÞTERÝEKLE.TextBox3.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 4).Value = MÜÞTERÝEKLE.TextBox4.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 5).Value = MÜÞTERÝEKLE.TextBox5.Text
Worksheets("MÜÞTERÝ").Range("b65536").End(xlUp).Offset(0, 6).Value = MÜÞTERÝEKLE.TextBox6.Text
Dim ds
Set ds = CreateObject("Scripting.FileSystemObject")
ds.CreateFolder ThisWorkbook.Path & "\TEKLÝFLER\" & MÜÞTERÝEKLE.TextBox1.Value
ds.CreateFolder ThisWorkbook.Path & "\FÝÞLER\" & MÜÞTERÝEKLE.TextBox1.Value
Unload MÜÞTERÝEKLE
müþterilistesi.Show 0
End If
End Sub
Private Sub CommandButton3_Click()
If MÜÞTERÝEKLE.TextBox1.Text = "" Then
MsgBox ("Lütfen Ürün Kodunu Giriniz...")
Else
Sheets("MÜÞTERÝ").Range("B" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox1.Text
Sheets("MÜÞTERÝ").Range("C" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox2.Text
Sheets("MÜÞTERÝ").Range("D" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox7.Text
Sheets("MÜÞTERÝ").Range("E" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox3.Text
Sheets("MÜÞTERÝ").Range("F" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox4.Text
Sheets("MÜÞTERÝ").Range("G" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox5.Text
Sheets("MÜÞTERÝ").Range("H" & müþterilistesi.ListBox1.ListIndex + 1).Value = MÜÞTERÝEKLE.TextBox6.Text
Unload MÜÞTERÝEKLE
End If
End Sub

Private Sub UserForm_Click()

End Sub
