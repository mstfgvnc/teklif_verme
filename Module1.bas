Attribute VB_Name = "Module1"
Sub auto_open()
Application.OnTime Now + TimeValue("00:00:02"), "kaydet"
End Sub
Sub kaydet()
If ThisWorkbook.Application.Visible = True Then
ThisWorkbook.Application.Visible = False
Else
End If 'buraya siz kendi istediðiniz kodlarý yazýnýz.
Call auto_open
End Sub
