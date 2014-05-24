VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================================================
'Name_ENG   Hash
'Name_CHS   哈希校验
'Version    API
'--------------------------------------------------------------------------------------
'源代码来自互联网
'======================================================================================

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phWndHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hWndHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hWndHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hWndHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long

Private Declare Function GetInputState Lib "user32" () As Long

Enum enum_HashAlgorithm
    [MD2] = 0
    [MD4] = 1
    [MD5] = 2
    [SHA1] = 3
End Enum

Public Algorithm As enum_HashAlgorithm

Dim hContext As Long
Dim hHash As Long

Private Sub Class_Initialize()
    Const Prov_RSA_Full = 1
    Const Crypt_NewKeySet = &H8
    
    Dim lngRes As Long
    lngRes = CryptAcquireContext(hContext, vbNullString, vbNullString, Prov_RSA_Full, 0)
    If lngRes = 0 Then
        If Err.LastDllError = &H80090016 Then
            lngRes = CryptAcquireContext(hContext, vbNullString, vbNullString, Prov_RSA_Full, Crypt_NewKeySet)
            If lngRes = 0 Then GoTo errInit
        Else
            GoTo errInit
        End If
    End If
    Algorithm = [MD5]
    Exit Sub
errInit:
    MsgBox "哈希校验模块初始化失败：无法创建套接字(Context)，请调用 Reset 方法重试。", 16
    Class_Terminate
End Sub
Private Sub Class_Terminate()
    If hContext Then
        CryptReleaseContext hContext, 0
        hContext = 0
    End If
End Sub
Sub Reset()
    Class_Initialize
End Sub

Sub Init(Alg As enum_HashAlgorithm)
    Const ALG_Class_Hash = 32768
    Const ALG_Type_Any = 0
    Const CALG = ALG_Class_Hash Or ALG_Type_Any
    Const ALG_SID_MD2 = 1
    Const ALG_SID_MD4 = 2
    Const ALG_SID_MD5 = 3
    Const ALG_SID_SHA1 = 4
    
    Dim AlgB As Long
    Select Case Alg
        Case MD2:  AlgB = CALG Or ALG_SID_MD2
        Case MD4:  AlgB = CALG Or ALG_SID_MD4
        Case MD5:  AlgB = CALG Or ALG_SID_MD5
        Case SHA1: AlgB = CALG Or ALG_SID_SHA1
    End Select
    CryptCreateHash hContext, AlgB, 0, 0, hHash
End Sub
Function Hash(Data() As Byte) As Boolean
    Hash = CryptHashData(hHash, Data(LBound(Data)), (UBound(Data) - LBound(Data) + 1), 0)
    If Hash = 0 Then
        MsgBox "哈希校验失败！", 16
        Terminate
        Exit Function
    End If
End Function
Function Over() As Byte()
    Const Hp_HashVal = 2
    Const Hp_HashSize = 4

    Dim valLen As Long, B() As Byte
    If CryptGetHashParam(hHash, Hp_HashSize, valLen, 4, 0) = 0 Then
        MsgBox "哈希校验失败：无法获取结果长度！", 16
        Terminate
        Exit Function
    End If
    ReDim B(1 To valLen)
    If CryptGetHashParam(hHash, Hp_HashVal, B(LBound(B)), valLen, 0) = 0 Then
        MsgBox "哈希校验失败：无法获取结果！", 16
        Terminate
        Exit Function
    End If
    Terminate
    Over = B
End Function
Private Sub Terminate()
    If hHash Then
        CryptDestroyHash hHash
        hHash = 0
    End If
End Sub

Rem 校验二进制数组
Function DataHash(Data() As Byte) As String
    Init Algorithm
    If Hash(Data) = True Then DataHash = BinToHex(Over)
End Function

Rem 校验文件
Function FileHashToBin(Path As String) As Byte()
    Const BlockSize As Long = 32 * 1024& '32 KB
    
On Error GoTo errFile
    If FileExist(Path) = False Then
        MsgBox "文件不存在！", 16
        Exit Function
    End If
    Init Algorithm
    Dim Fn As Integer, FL As Long, B() As Byte, i&, j&
    Fn = FreeFile
    Open Path For Binary As #Fn
    FL = FileLen(Path)
    ReDim B(1 To BlockSize)
    For i = 1 To FL \ BlockSize
        Get #Fn, , B
        If Hash(B) = False Then GoTo errFile
        DoEvents
    Next i
    FL = FL Mod BlockSize
    If FL Then
        ReDim ABHash(1 To FL)
        Get #Fn, , B
        If Hash(B) = False Then GoTo errFile
    End If
    Close #Fn
    FileHashToBin = Over
    Exit Function
errFile:
    MsgBox Err.Number & ":" & Err.Description, 16
    Close #Fn
    Terminate
End Function
Function FileHash(Path As String) As String
    FileHash = BinToHex(FileHashToBin(Path))
End Function
Private Function FileExist(Path As String) As Boolean
    Path = CorrectPath(Path)
    Dim tN As String
    tN = Right$(Path, Len(Path) - InStrRev(Path, "\"))
    Dim tS As String
    tS = Dir(Path, vbHidden Or vbReadOnly Or vbSystem Or vbNormal)
    FileExist = (tS = tN)
End Function
Private Function CorrectPath(ByVal Path As String) As String
    Path = Replace(Path, "//", "/")
    Path = Replace(Path, "\\", "\")
    Path = Replace(Path, """", "")
    CorrectPath = Path
End Function

Rem 校验文本
Function TextHash(Text As String) As String
    TextHash = DataHash(StrConv(Text, vbFromUnicode))
End Function
Private Function BinToHex(B() As Byte) As String
    On Error GoTo errB2H
    Dim H() As String
    Dim BS As Long, BL As Long, BU As Long
    BL = LBound(B): BU = UBound(B): BS = BU - BL + 1
    If BS Then
        ReDim H(BL To BU)
        Dim i&
        For i = BL To BU
            H(i) = Right$("0" & Hex$(B(i)), 2)
            If GetInputState Then DoEvents
        Next i
        BinToHex = Join(H, vbNullString)
    End If
errB2H:
End Function
Private Function HexToBin(H As String, Optional lngLBound As Long) As Byte()
    On Error GoTo errH2B
    Dim B() As Byte
    Dim BS As Long
    BS = Len(H)
    If BS Then
        If BS Mod 2 Then
            ReDim B(lngLBound To BS \ 2 + lngLBound)
            B(UBound(B)) = Val("&H" & Right(H, 1))
        Else
            ReDim B(lngLBound To BS \ 2 + lngLBound - 1)
        End If
        Dim i&, j&
        j = lngLBound
        For i = 1 To BS Step 2
            B(j) = Val("&H" & Mid(H, i, 2))
            j = j + 1
            If GetInputState Then DoEvents
        Next i
        HexToBin = B
    End If
errH2B:
End Function
