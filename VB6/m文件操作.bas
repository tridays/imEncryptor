Attribute VB_Name = "m文件操作"
Option Explicit

'======================================================================================
'Name_ENG   File Handle Module
'Name_CHS   文件操作模块
'Version    1.5
'Author     tridays
'Studio     (none)
'Date       12-7-25
'--------------------------------------------------------------------------------------
'感谢使用,请保留版权信息
'======================================================================================

'读取二进制文件（路径, 起始位置, 读取长度, 返回数组的下标上界）
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
'写入二进制文件（数组, 路径, 起始位置, 保存前是否清空【默认是】）
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

'快速读取文本文档
Function LoadTxt(Path As String, Optional Position As Long = 1, Optional Length As Long) As String
    LoadTxt = StrConv(LoadBin(Path, Position, Length), vbUnicode)
End Function
'保存文本
Sub SaveTxt(Text As String, Path As String, Optional Position As Long = 1, Optional Empty_at_First As Boolean = True)
    SaveBin StrConv(Text, vbFromUnicode), Path, Position, Empty_at_First
End Sub

'检测文件状态
Function CheckFile(Path As String) As Long
    If Path = "" Then Exit Function
    On Error GoTo NoFile
    CorrectPath Path
    FileLen Path
    CheckFile = 0                   '文件存在
    Exit Function
NoFile:
    On Error GoTo errPath
    Dim Fn As Integer
    Fn = FreeFile
    Open Path For Output As #Fn
    Close #Fn
    Kill Path
    CheckFile = 1                   '文件不存在，但路径正确
    Exit Function
errPath:
    CheckFile = 2                   '路径错误
End Function

'检测文件是否存在
Function FileExist(Path As String) As Boolean
    Path = CorrectPath(Path)
    Dim tN As String
    tN = Right$(Path, Len(Path) - InStrRev(Path, "\"))
    Dim tS As String
    tS = Dir(Path, vbHidden Or vbReadOnly Or vbSystem Or vbNormal)
    FileExist = (tS = tN)
End Function

'检测目录是否存在
Function FolderExist(Path As String) As Boolean
    Path = CorrectPath(Path)
    If Dir(Path, vbDirectory) Then
        FolderExist = (GetAttr(Path) And vbDirectory)
    End If
End Function

'随机创建临时文件(返回绝对路径)
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

'将以Byte计数的文件长度换成更高单位(保留两位小数)
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

'纠正部分路径的语法错误(对于VB来说)
Function CorrectPath(ByVal Path As String) As String
    Path = Replace(Path, "//", "/")
    Path = Replace(Path, "\\", "\")
    Path = Replace(Path, """", "")
    CorrectPath = Path
End Function

'后台注册组件(从资源文件)
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
'            RegDLL = 2                                                      '注册失败
'            Exit Function
'          End If
'        End If
'        RegDLL = 0                                                          '注册成功
'    Else
'        RegDLL = 1                                                          '已注册
'    End If
'errRegDLL:
'End Function
