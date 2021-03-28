VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VTEKLÝFLER 
   Caption         =   "VERÝLEN TEKLÝFLER"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   OleObjectBlob   =   "VTEKLÝFLER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VTEKLÝFLER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
X = Sheets("tekliftablosu").Range("B" & ListBox1.ListIndex + 2).Value
y = Sheets("tekliftablosu").Range("c" & ListBox1.ListIndex + 2).Value
CreateObject("Shell.Application").Open ThisWorkbook.Path & "\TEKLÝFLER\" & X & "\" & y & ".xls"
End Sub
Private Sub UserForm_Initialize()
ListBox1.Clear
ListBox1.ColumnCount = 2
ListBox1.ColumnWidths = "100;100"
ListBox1.RowSource = "tekliftablosu!b2:c" & Sheets("tekliftablosu").Range("c65656").End(xlUp).Row
End Sub
