VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VF��LER 
   Caption         =   "KES�LEN F��LER"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   OleObjectBlob   =   "VF��LER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "VF��LER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
X = Sheets("fi�tablosu").Range("B" & ListBox1.ListIndex + 2).Value
y = Sheets("fi�tablosu").Range("c" & ListBox1.ListIndex + 2).Value
CreateObject("Shell.Application").Open ThisWorkbook.Path & "\F��LER\" & X & "\" & y & ".xls"
End Sub
Private Sub UserForm_Initialize()
ListBox1.Clear
ListBox1.ColumnCount = 2
ListBox1.ColumnWidths = "100;100"
ListBox1.RowSource = "fi�tablosu!b2:c" & Sheets("fi�tablosu").Range("c65656").End(xlUp).Row
End Sub
