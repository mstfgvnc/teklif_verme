VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} F�� 
   Caption         =   "F�� KES"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11145
   OleObjectBlob   =   "F��.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "F��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
m��terilistesi.CommandButton1.Enabled = False
m��terilistesi.CommandButton2.Enabled = False
m��terilistesi.CommandButton3.Enabled = False
m��terilistesi.CommandButton4.Enabled = False
m��terilistesi.Show 0
End Sub
Private Sub CommandButton2_Click()
UserForm3.CommandButton1.Enabled = False
UserForm3.CommandButton3.Enabled = False
UserForm3.CommandButton4.Enabled = False
UserForm3.CommandButton5.Enabled = False
F��.Hide
UserForm3.Show 0
End Sub
Private Sub CommandButton3_Click()
answer = MsgBox(Sheets("aaa").Range("E" & ListBox1.ListIndex + 21).Value & " adl� �r�n� tekliften ��karmak istedi�inize emin misiniz?", vbYesNo + vbQuestion, "M��TER� L�STES�")
If answer = vbYes Then
Dim i, a, b As Integer
a = Sheets("bbb").Range("E" & ListBox1.ListIndex + 6).Column
b = Sheets("bbb").Range("E" & ListBox1.ListIndex + 6).Row
Sheets("bbb").Range("D" & ListBox1.ListIndex + 5).ClearContents
Sheets("bbb").Range("E" & ListBox1.ListIndex + 5).ClearContents
Sheets("bbb").Range("C" & ListBox1.ListIndex + 5).ClearContents
If Sheets("bbb").Range("E" & ListBox1.ListIndex + 6).Value = "" Then
Else
i = Worksheets("bbb").Range("e31").End(xlUp).Row
Sheets("bbb").Select
Sheets("bbb").Range(Cells(b, a - 2), Cells(i, a)).Select
Selection.Cut
ActiveSheet.Cells(b - 1, a - 2).Select
ActiveSheet.Paste
Sheets("bbb").Select
Sheets("bbb").Range(Cells(30, 3), Cells(30, 5)).Select
Selection.Copy
ActiveSheet.Cells(i, 3).Select
ActiveSheet.Paste
End If
Else
End If
End Sub
Private Sub CommandButton5_Click()
answer = MsgBox(" Fi�i kaydetmek istedi�inize emin misiniz?", vbYesNo + vbQuestion, "F�� KES")
If answer = vbYes Then
a = Format(Now, "dd.mm.yyyy-hh.mm")
b = Left(a, 10)
C = Right(a, 5)
Sheets("bbb").Range("d4").Value = F��.TextBox1.Value
Sheets("bbb").Range("f3").Value = b
Sheets("bbb").Range("c33").Value = F��.TextBox3.Value
Sheets("bbb").Range("e33").Value = F��.TextBox4.Value
ActiveSheet.Copy
ActiveSheet.SaveAs ThisWorkbook.Path & "\F��LER\" & F��.TextBox1.Text & "\" & b & " " & C & ".xls"
Workbooks("TEKL�F MG.xlsm").Worksheets("fi�tablosu").Activate
Worksheets("fi�tablosu").Range("b65536").End(xlUp).Offset(1, 0).Value = F��.TextBox1.Value
Worksheets("fi�tablosu").Range("b65536").End(xlUp).Offset(0, 1).Value = b & " " & C
Unload F��
Workbooks(b & " " & C & ".xls").Close
VF��LER.Show 0
Else
End If
End Sub
Private Sub UserForm_Initialize()
Dim ts
Set ts = ActiveSheet
a = ActiveSheet.Name
ListBox1.Clear
ListBox1.ColumnCount = 4
ListBox1.ColumnWidths = "30;80;100;100"
ListBox1.RowSource = a & "!B5:E" & ts.Range("E31").End(xlUp).Row
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error Resume Next
Application.DisplayAlerts = False
Sheets("bbb").Delete
End Sub

