VERSION 5.00
Begin VB.Form Passwrd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter the Code"
   ClientHeight    =   1620
   ClientLeft      =   6015
   ClientTop       =   6240
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   3960
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtcd 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   330
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Enter the CODE:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Passwrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Hide
End Sub

Private Sub Command2_Click()
txtcd = ""
Hide
End Sub
