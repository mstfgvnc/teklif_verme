VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} m��terilistesi 
   Caption         =   "M��TER� L�STES�"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690
   OleObjectBlob   =   "m��terilistesi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "m��terilistesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
M��TER�EKLE.CommandButton3.Enabled = False
M��TER�EKLE.CommandButton4.Enabled = False
Unload m��terilistesi
M��TER�EKLE.Show 0
End Sub
Private Sub CommandButton2_Click()
answer = MsgBox(Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 1).Value & " firmay� silmek istedi�inize emin misiniz?", vbYesNo + vbQuestion, "M��TER� L�STES�")
If answer = vbYes Then
Dim i, a, b As Integer
a = Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 2).Column
b = Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 2).Row
Sheets("M��TER�").Range("c" & ListBox1.ListIndex + 1).ClearContents
Sheets("M��TER�").Range("d" & ListBox1.ListIndex + 1).ClearContents
Sheets("M��TER�").Range("e" & ListBox1.ListIndex + 1).ClearContents
Sheets("M��TER�").Range("f" & ListBox1.ListIndex + 1).ClearContents
Sheets("M��TER�").Range("h" & ListBox1.ListIndex + 1).ClearContents
Sheets("M��TER�").Range("g" & ListBox1.ListIndex + 1).ClearContents
Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 1).ClearContents
If Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 2).Value = "" Then
Else
i = Worksheets("M��TER�").Range("b655336").End(xlUp).Row
Sheets("M��TER�").Select
Sheets("M��TER�").Range(Cells(b, a), Cells(i, a + 7)).Select
Selection.Cut
ActiveSheet.Cells(b - 1, a).Select
ActiveSheet.Paste
End If
Else
End If
End Sub
Private Sub CommandButton3_Click()
M��TER�EKLE.CommandButton2.Enabled = False
M��TER�EKLE.CommandButton4.Enabled = False
On Error Resume Next
M��TER�EKLE.TextBox1.Text = Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.TextBox2.Text = Sheets("M��TER�").Range("C" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.TextBox7.Text = Sheets("M��TER�").Range("D" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.TextBox3.Text = Sheets("M��TER�").Range("E" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.TextBox4.Text = Sheets("M��TER�").Range("F" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.TextBox5.Text = Sheets("M��TER�").Range("G" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.TextBox6.Text = Sheets("M��TER�").Range("H" & ListBox1.ListIndex + 1).Value
M��TER�EKLE.Show 0
End Sub
Private Sub CommandButton4_Click()
On Error Resume Next
TEKL�F.TextBox1.Text = Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 1).Value
TEKL�F.TextBox2.Text = Sheets("M��TER�").Range("C" & ListBox1.ListIndex + 1).Value
TEKL�F.TextBox7.Text = Sheets("M��TER�").Range("D" & ListBox1.ListIndex + 1).Value
TEKL�F.TextBox3.Text = Sheets("M��TER�").Range("E" & ListBox1.ListIndex + 1).Value
TEKL�F.TextBox4.Text = Sheets("M��TER�").Range("F" & ListBox1.ListIndex + 1).Value
TEKL�F.TextBox5.Text = Sheets("M��TER�").Range("G" & ListBox1.ListIndex + 1).Value
TEKL�F.TextBox6.Text = Sheets("M��TER�").Range("H" & ListBox1.ListIndex + 1).Value
Unload m��terilistesi
End Sub
Private Sub CommandButton5_Click()
F��.TextBox1.Text = Sheets("M��TER�").Range("B" & ListBox1.ListIndex + 1).Value
F��.TextBox2.Text = Sheets("M��TER�").Range("C" & ListBox1.ListIndex + 1).Value
Unload m��terilistesi
End Sub

Private Sub CommandButton6_Click()

End Sub

Private Sub UserForm_Initialize()
Dim ts
Set ts = Sheets("M��TER�")
ListBox1.Clear
ListBox1.ColumnCount = 8
ListBox1.ColumnWidths = "20;150;80;150;40;40;40;40"
ListBox1.RowSource = "M��TER�!A1:H" & ts.Range("B" & Rows.Count).End(xlUp).Row
End Sub
