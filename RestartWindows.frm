VERSION 5.00
Begin VB.Form RestartWindows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   2160
   ClientLeft      =   5715
   ClientTop       =   5850
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4110
   Begin VB.CommandButton cmdOtm 
      Caption         =   "Abrogation"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Вы обязательно должны перезапустить Windows. Сохраните все данные и нажмите кнопку Restart."
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "RestartWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_LOGOFF = 0
Const EWX_FORCE = 4
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Sub cmdOtm_Click()
Unload RestartWindows
End Sub
Private Sub cmdRestart_Click()
' перезапускаем Windows (срабатывает в Windows 95/NT)
ExitWindowsEx EWX_LOGOFF, 0
End Sub

