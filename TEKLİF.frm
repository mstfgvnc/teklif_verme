VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TEKL�F 
   Caption         =   "TEKL�F HAZIRLAMA"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15270
   OleObjectBlob   =   "TEKL�F.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TEKL�F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
m��terilistesi.CommandButton1.Enabled = False
m��terilistesi.CommandButton2.Enabled = False
m��terilistesi.CommandButton3.Enabled = False
m��terilistesi.CommandButton5.Enabled = False
m��terilistesi.Show 0
End Sub
Private Sub CommandButton2_Click()
UserForm3.CommandButton1.Enabled = False
UserForm3.CommandButton3.Enabled = False
UserForm3.CommandButton4.Enabled = False
UserForm3.CommandButton6.Enabled = False
TEKL�F.Hide
UserForm3.Show 0
End Sub
Private Sub CommandButton3_Click()
answer = MsgBox(Sheets("aaa").Range("E" & ListBox1.ListIndex + 21).Value & " adl� �r�n� tekliften ��karmak istedi�inize emin misiniz?", vbYesNo + vbQuestion, "M��TER� L�STES�")
If answer = vbYes Then
Dim i, a, b As Integer
a = Sheets("aaa").Range("E" & ListBox1.ListIndex + 22).Column
b = Sheets("aaa").Range("E" & ListBox1.ListIndex + 22).Row
Sheets("aaa").Range("D" & ListBox1.ListIndex + 21).ClearContents
Sheets("aaa").Range("E" & ListBox1.ListIndex + 21).ClearContents
Sheets("aaa").Range("F" & ListBox1.ListIndex + 21).ClearContents
Sheets("aaa").Range("G" & ListBox1.ListIndex + 21).ClearContents
Sheets("aaa").Range("H" & ListBox1.ListIndex + 21).ClearContents
Sheets("aaa").Range("I" & ListBox1.ListIndex + 21).ClearContents
If Sheets("aaa").Range("d" & ListBox1.ListIndex + 22).Value = "" Then
Else
i = Worksheets("aaa").Range("D48").End(xlUp).Row
Sheets("aaa").Select
Sheets("aaa").Range(Cells(b, a - 1), Cells(i, a + 5)).Select
Selection.Cut
ActiveSheet.Cells(b - 1, a - 1).Select
ActiveSheet.Paste
Sheets("aaa").Select
Sheets("aaa").Range(Cells(47, 4), Cells(47, 9)).Select
Selection.Copy
ActiveSheet.Cells(i, 4).Select
ActiveSheet.Paste
End If
Else
End If
End Sub
Private Sub CommandButton4_Click()
a = InputBox("L�tfen �skonto Oran�n� TL baz�nda Giriniz...")
Sheets("aaa").Range("j49").Value = a
End Sub
Private Sub CommandButton5_Click()
answer = MsgBox(" Teklifi kaydetmek istedi�inize emin misiniz?", vbYesNo + vbQuestion, "M��TER� L�STES�")
If answer = vbYes Then
a = Format(Now, "dd.mm.yyyy-hh.mm")
b = Left(a, 10)
C = Right(a, 5)
Sheets("aaa").Range("d13").Value = TEKL�F.TextBox1.Value
Sheets("aaa").Range("d14").Value = TEKL�F.TextBox2.Value
Sheets("aaa").Range("d15").Value = TEKL�F.TextBox7.Value
Sheets("aaa").Range("d17").Value = TEKL�F.TextBox3.Value
Sheets("aaa").Range("d18").Value = TEKL�F.TextBox4.Value
Sheets("aaa").Range("�13").Value = a
Sheets("aaa").Range("�14").Value = TEKL�F.TextBox8.Value
Sheets("aaa").Range("�17").Value = TEKL�F.TextBox9.Value
Sheets("aaa").Range("e49").Value = TEKL�F.TextBox10.Value
Sheets("aaa").Range("d56").Value = TEKL�F.TextBox11.Value
ActiveSheet.Copy
ActiveSheet.SaveAs ThisWorkbook.Path & "\TEKL�FLER\" & TEKL�F.TextBox1.Text & "\" & b & " " & C & ".xls"
Workbooks("TEKL�F MG.xlsm").Worksheets("tekliftablosu").Activate
Worksheets("tekliftablosu").Range("b65536").End(xlUp).Offset(1, 0).Value = TEKL�F.TextBox1.Value
Worksheets("tekliftablosu").Range("b65536").End(xlUp).Offset(0, 1).Value = b & " " & C
Unload TEKL�F
Workbooks(b & " " & C & ".xls").Close
VTEKL�FLER.Show 0
Else
End If
End Sub
Private Sub UserForm_Initialize()
Dim ts
Set ts = ActiveSheet
a = ActiveSheet.Name
ListBox1.Clear
ListBox1.ColumnCount = 8
ListBox1.ColumnWidths = "30;80;100;100;50;50;50;50"
ListBox1.RowSource = a & "!C21:j" & ts.Range("d48").End(xlUp).Row
ListBox2.RowSource = "aaa!j48"
ListBox3.RowSource = "aaa!j51"
ListBox4.RowSource = "aaa!j50"
ListBox5.RowSource = "aaa!j49"
ListBox6.RowSource = "aaa!j52"
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error Resume Next
Application.DisplayAlerts = False
Sheets("aaa").Delete
End Sub
