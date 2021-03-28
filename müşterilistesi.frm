VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} müþterilistesi 
   Caption         =   "MÜÞTERÝ LÝSTESÝ"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690
   OleObjectBlob   =   "müþterilistesi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "müþterilistesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
MÜÞTERÝEKLE.CommandButton3.Enabled = False
MÜÞTERÝEKLE.CommandButton4.Enabled = False
Unload müþterilistesi
MÜÞTERÝEKLE.Show 0
End Sub
Private Sub CommandButton2_Click()
answer = MsgBox(Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value & " firmayý silmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "MÜÞTERÝ LÝSTESÝ")
If answer = vbYes Then
Dim i, a, b As Integer
a = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 2).Column
b = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 2).Row
Sheets("MÜÞTERÝ").Range("c" & ListBox1.ListIndex + 1).ClearContents
Sheets("MÜÞTERÝ").Range("d" & ListBox1.ListIndex + 1).ClearContents
Sheets("MÜÞTERÝ").Range("e" & ListBox1.ListIndex + 1).ClearContents
Sheets("MÜÞTERÝ").Range("f" & ListBox1.ListIndex + 1).ClearContents
Sheets("MÜÞTERÝ").Range("h" & ListBox1.ListIndex + 1).ClearContents
Sheets("MÜÞTERÝ").Range("g" & ListBox1.ListIndex + 1).ClearContents
Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).ClearContents
If Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 2).Value = "" Then
Else
i = Worksheets("MÜÞTERÝ").Range("b655336").End(xlUp).Row
Sheets("MÜÞTERÝ").Select
Sheets("MÜÞTERÝ").Range(Cells(b, a), Cells(i, a + 7)).Select
Selection.Cut
ActiveSheet.Cells(b - 1, a).Select
ActiveSheet.Paste
End If
Else
End If
End Sub
Private Sub CommandButton3_Click()
MÜÞTERÝEKLE.CommandButton2.Enabled = False
MÜÞTERÝEKLE.CommandButton4.Enabled = False
On Error Resume Next
MÜÞTERÝEKLE.TextBox1.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.TextBox2.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.TextBox7.Text = Sheets("MÜÞTERÝ").Range("D" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.TextBox3.Text = Sheets("MÜÞTERÝ").Range("E" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.TextBox4.Text = Sheets("MÜÞTERÝ").Range("F" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.TextBox5.Text = Sheets("MÜÞTERÝ").Range("G" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.TextBox6.Text = Sheets("MÜÞTERÝ").Range("H" & ListBox1.ListIndex + 1).Value
MÜÞTERÝEKLE.Show 0
End Sub
Private Sub CommandButton4_Click()
On Error Resume Next
TEKLÝF.TextBox1.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
TEKLÝF.TextBox2.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
TEKLÝF.TextBox7.Text = Sheets("MÜÞTERÝ").Range("D" & ListBox1.ListIndex + 1).Value
TEKLÝF.TextBox3.Text = Sheets("MÜÞTERÝ").Range("E" & ListBox1.ListIndex + 1).Value
TEKLÝF.TextBox4.Text = Sheets("MÜÞTERÝ").Range("F" & ListBox1.ListIndex + 1).Value
TEKLÝF.TextBox5.Text = Sheets("MÜÞTERÝ").Range("G" & ListBox1.ListIndex + 1).Value
TEKLÝF.TextBox6.Text = Sheets("MÜÞTERÝ").Range("H" & ListBox1.ListIndex + 1).Value
Unload müþterilistesi
End Sub
Private Sub CommandButton5_Click()
FÝÞ.TextBox1.Text = Sheets("MÜÞTERÝ").Range("B" & ListBox1.ListIndex + 1).Value
FÝÞ.TextBox2.Text = Sheets("MÜÞTERÝ").Range("C" & ListBox1.ListIndex + 1).Value
Unload müþterilistesi
End Sub

Private Sub CommandButton6_Click()

End Sub

Private Sub UserForm_Initialize()
Dim ts
Set ts = Sheets("MÜÞTERÝ")
ListBox1.Clear
ListBox1.ColumnCount = 8
ListBox1.ColumnWidths = "20;150;80;150;40;40;40;40"
ListBox1.RowSource = "MÜÞTERÝ!A1:H" & ts.Range("B" & Rows.Count).End(xlUp).Row
End Sub
