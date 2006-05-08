VERSION 5.00
Begin VB.Form Informasion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information"
   ClientHeight    =   1605
   ClientLeft      =   5040
   ClientTop       =   5070
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5325
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   360
      Top             =   1080
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   240
      Picture         =   "Informasion.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "“ы сюда не ходи, ты туда ходи, ато кирпич башка упадет, совсем мертвый будешь."
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Informasion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Informasion
End Sub
Private Sub Timer1_Timer()
Unload Informasion
End Sub
