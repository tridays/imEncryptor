Attribute VB_Name = "m�ļ�����"
Option Explicit

'======================================================================================
'Name_ENG   File Handle Module
'Name_CHS   �ļ�����ģ��
'Version    1.5
'Author     tridays
'Studio     (none)
'Date       12-7-25
'--------------------------------------------------------------------------------------
'��лʹ��,�뱣����Ȩ��Ϣ
'======================================================================================

'��ȡ�������ļ���·��, ��ʼλ��, ��ȡ����, ����������±��Ͻ磩
Function LoadBin(Path As String, Optional Position As Long = 1, Optional Length As Long, Optional lngLBound As Long) As Byte()
    Dim Fn As Integer
    Fn = FreeFile
    Open Path For Binary Access Read As #Fn
    If Length = 0 Then
        ReDim LoadBin(lngLBound To LOF(Fn) + lngLBound - 1)
    Else
        ReDim LoadBin(lngLBound To Length + lngLBound - 1)
    End If
    Get #Fn, Position, LoadBin
    Close #Fn
End Function
'д��������ļ�������, ·��, ��ʼλ��, ����ǰ�Ƿ���ա�Ĭ���ǡ���
Sub SaveBin(Data() As Byte, Path As String, Optional Position As Long = 1, Optional Empty_at_First As Boolean = True)
    Dim Fn As Integer
    Fn = FreeFile
    If Empty_at_First Then
        Open Path For Output As #Fn
        Close #Fn
    End If
    Open Path For Binary Access Write As #Fn
    Put #Fn, Position, Data
    Close #Fn
End Sub

'���ٶ�ȡ�ı��ĵ�
Function LoadTxt(Path As String, Optional Position As Long = 1, Optional Length As Long) As String
    LoadTxt = StrConv(LoadBin(Path, Position, Length), vbUnicode)
End Function
'�����ı�
Sub SaveTxt(Text As String, Path As String, Optional Position As Long = 1, Optional Empty_at_First As Boolean = True)
    SaveBin StrConv(Text, vbFromUnicode), Path, Position, Empty_at_First
End Sub

'����ļ�״̬
Function CheckFile(Path As String) As Long
    If Path = "" Then Exit Function
    On Error GoTo NoFile
    CorrectPath Path
    FileLen Path
    CheckFile = 0                   '�ļ�����
    Exit Function
NoFile:
    On Error GoTo errPath
    Dim Fn As Integer
    Fn = FreeFile
    Open Path For Output As #Fn
    Close #Fn
    Kill Path
    CheckFile = 1                   '�ļ������ڣ���·����ȷ
    Exit Function
errPath:
    CheckFile = 2                   '·������
End Function

'����ļ��Ƿ����
Function FileExist(Path As String) As Boolean
    Path = CorrectPath(Path)
    Dim tN As String
    tN = Right$(Path, Len(Path) - InStrRev(Path, "\"))
    Dim tS As String
    tS = Dir(Path, vbHidden Or vbReadOnly Or vbSystem Or vbNormal)
    FileExist = (tS = tN)
End Function

'���Ŀ¼�Ƿ����
Function FolderExist(Path As String) As Boolean
    Path = CorrectPath(Path)
    If Dir(Path, vbDirectory) Then
        FolderExist = (GetAttr(Path) And vbDirectory)
    End If
End Function

'���������ʱ�ļ�(���ؾ���·��)
Function GetTempPath() As String
On Error GoTo ErrFso
    Dim Fso As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    GetTempPath = Replace(Fso.GetSpecialFolder(2) & "\" & Fso.GetTempName, "\\", "\")
    Do Until Fso.FileExists(GetTempPath) = False
      GetTempPath = Replace(Fso.GetSpecialFolder(2) & "\" & Fso.GetTempName, "\\", "\")
      DoEvents
    Loop
ErrFso:
    Set Fso = Nothing
End Function

'����Byte�������ļ����Ȼ��ɸ��ߵ�λ(������λС��)
Function GetFileLen(Path As String) As String
    GetFileLen = GetFileLenFromValue(FileLen(Path))
End Function
Function GetFileLenFromValue(Value) As String
    Dim Level As Integer
    Dim tV As Double
    Dim tS As String
    tV = Value
    Do While tV >= 1024
        tV = tV / 1024
        Level = Level + 1
    Loop
    Select Case Level
    Case 0:     tS = " Byte(s)"
    Case 1:     tS = " KB"
    Case 2:     tS = " MB"
    Case 3:     tS = " GB"
    Case 4:     tS = " TB"
    Case Else:  tS = " More than TB"
    End Select
    GetFileLenFromValue = Format(tV, "0.00") & tS
End Function

'��������·�����﷨����(����VB��˵)
Function CorrectPath(ByVal Path As String) As String
    Path = Replace(Path, "//", "/")
    Path = Replace(Path, "\\", "\")
    Path = Replace(Path, """", "")
    CorrectPath = Path
End Function

'��̨ע�����(����Դ�ļ�)
'Function RegDLL(Number As Long, Types As String, FileName As String, Optional Regist As Boolean = False) As Long
'    On Error GoTo errRegDLL
'    Dim tPath As String
'    tPath = Environ("Windir") & "\System32\" & FileName
'    CorrectPath tPath
'    If FileExist(tPath) = 1 Then
'        Dim tB() As Byte
'        tB = LoadResData(Number, Types)
'        SaveBin tB, tPath
'        If Regist = True Then
'          RegDLL = Shell("Regsvr32 /s " & FileName, vbMinimizedNoFocus)
'          If RegDLL = 0 Then
'            RegDLL = 2                                                      'ע��ʧ��
'            Exit Function
'          End If
'        End If
'        RegDLL = 0                                                          'ע��ɹ�
'    Else
'        RegDLL = 1                                                          '��ע��
'    End If
'errRegDLL:
'End Function
