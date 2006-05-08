VERSION 5.00
Begin VB.Form StepTwo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ruin Killer"
   ClientHeight    =   1665
   ClientLeft      =   4995
   ClientTop       =   4845
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3450
   Begin VB.CommandButton Command1 
      Caption         =   "Push"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Killing: step two"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "StepTwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal IpBuffer As String, ByVal nSize As Long) As Long
Dim dirc As String
Dim strFileName As String
Private Sub Command1_Click()
Dim IngRtn As Long
Const MAXWINPATH = 144
dirc = Space(MAXWINPATH)
IngRtn = GetWindowsDirectory(dirc, MAXWINPATH)
dirc = Left(dirc, IngRtn)  ' отсекаем лишнее
strFileName = Dir(dirc & "\system32\")
Do While strFileName <> ""
If strFileName = "DateForBegin" Then Kill dirc & "\system32\DateForBegin"
strFileName = Dir
Loop
Unload StepTwo
End Sub

