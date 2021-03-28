VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FÝÞ 
   Caption         =   "FÝÞ KES"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11145
   OleObjectBlob   =   "FÝÞ.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FÝÞ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
müþterilistesi.CommandButton1.Enabled = False
müþterilistesi.CommandButton2.Enabled = False
müþterilistesi.CommandButton3.Enabled = False
müþterilistesi.CommandButton4.Enabled = False
müþterilistesi.Show 0
End Sub
Private Sub CommandButton2_Click()
UserForm3.CommandButton1.Enabled = False
UserForm3.CommandButton3.Enabled = False
UserForm3.CommandButton4.Enabled = False
UserForm3.CommandButton5.Enabled = False
FÝÞ.Hide
UserForm3.Show 0
End Sub
Private Sub CommandButton3_Click()
answer = MsgBox(Sheets("aaa").Range("E" & ListBox1.ListIndex + 21).Value & " adlý ürünü tekliften çýkarmak istediðinize emin misiniz?", vbYesNo + vbQuestion, "MÜÞTERÝ LÝSTESÝ")
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
answer = MsgBox(" Fiþi kaydetmek istediðinize emin misiniz?", vbYesNo + vbQuestion, "FÝÞ KES")
If answer = vbYes Then
a = Format(Now, "dd.mm.yyyy-hh.mm")
b = Left(a, 10)
C = Right(a, 5)
Sheets("bbb").Range("d4").Value = FÝÞ.TextBox1.Value
Sheets("bbb").Range("f3").Value = b
Sheets("bbb").Range("c33").Value = FÝÞ.TextBox3.Value
Sheets("bbb").Range("e33").Value = FÝÞ.TextBox4.Value
ActiveSheet.Copy
ActiveSheet.SaveAs ThisWorkbook.Path & "\FÝÞLER\" & FÝÞ.TextBox1.Text & "\" & b & " " & C & ".xls"
Workbooks("TEKLÝF MG.xlsm").Worksheets("fiþtablosu").Activate
Worksheets("fiþtablosu").Range("b65536").End(xlUp).Offset(1, 0).Value = FÝÞ.TextBox1.Value
Worksheets("fiþtablosu").Range("b65536").End(xlUp).Offset(0, 1).Value = b & " " & C
Unload FÝÞ
Workbooks(b & " " & C & ".xls").Close
VFÝÞLER.Show 0
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

