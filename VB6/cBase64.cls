VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================================================
'Name_ENG   Base64 Encoder
'Name_CHS   Base64 编码器
'Version    1.3
'Author     Tridays
'Studio     ImageSoft
'Date       13-10-15
'--------------------------------------------------------------------------------------
'感谢使用,请保留版权信息
'======================================================================================

'=================================
'ConvList 提供显示字符的查看、修改
'Reserved 提供对保留字的查看、修改
'=================================

Const Bit72 = 252   '128 Or 64 Or 32 Or 16 Or 8 Or 4
Const Bit10 = 3     '2 Or 1
Const Bit74 = 240   '128 Or 64 Or 32 Or 16
Const Bit30 = 15    '8 Or 4 Or 2 Or 1
Const Bit76 = 192   '128 Or 64
Const Bit50 = 63    '32 Or 16 Or 8 Or 4 Or 2 Or 1

Const Bit54 = 48    '32 Or 16
Const Bit52 = 61    '32 Or 16 Or 8 Or 4

Private Declare Function GetInputState Lib "user32" () As Long

Event Init(Duration As Long)
Event Progress(Position As Long)
Event Over()
Event Error(Number As Integer, Description As String)

'Private Declare Sub CpyMem4 Lib "msvbvm60.dll" Alias "GetMem4" (ByRef RefSrc As Any, ByRef RefDst As Any)

Enum enum_Base64_InitList
    [B64_Normal] = 0
    [B64_URL]
    [B64_RegEx]
    [B64_Program1]
    [B64_Program2]
    [B64_XML_Nmtoken]
    [B64_XML_Name]
End Enum

Dim BaseStr(0 To 64) As Byte '第64号是保留字 the No.64 is Reserved.
Dim BaseRev(0 To 255) As Byte

Private Sub Class_Initialize()
    Initialize 0
End Sub
Sub Initialize(Optional InitList As enum_Base64_InitList = B64_Normal)
    Dim i As Byte, tV As Byte
    '标准字符 62个
    Const InitStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    For i = 1 To Len(InitStr)
        SetBaseStr i - 1, Asc(Mid(InitStr, i, 1))
    Next i
    '特殊字符 3个（含保留字）
    Select Case InitList
    Case 0 '标准
        SetBaseStr 62, 43 '+
        SetBaseStr 63, 47 '/
        SetBaseStr 64, 61 '=
    Case 1 'URL
        SetBaseStr 62, 42 '*
        SetBaseStr 63, 45 '-
        SetBaseStr 64, 61 '=
    Case 2 '正则表达式
        SetBaseStr 62, 33 '+
        SetBaseStr 63, 45 '-
        SetBaseStr 64, 61 '=
    Case 3 '编程 1
        SetBaseStr 62, 95 '_
        SetBaseStr 63, 45 '-
        SetBaseStr 64, 61 '=
    Case 4 '编程 2
        SetBaseStr 62, 46 '.
        SetBaseStr 63, 95 '_
        SetBaseStr 64, 61 '=
    Case 5 'XML Nmtoken
        SetBaseStr 62, 46 '.
        SetBaseStr 63, 45 '-
        SetBaseStr 64, 61 '=
    Case 6 'XML Name
        SetBaseStr 62, 95 '_
        SetBaseStr 63, 59 ':
        SetBaseStr 64, 61 '=
    End Select
End Sub
Private Sub SetBaseStr(Pos As Byte, Value As Byte)
    BaseStr(Pos) = Value
    BaseRev(Value) = Pos
    'Debug.Print Chr(Value);
End Sub

Property Get ConvList(Index As Byte) As Byte
    Select Case Index
    Case Is < 0, Is > 64
        'none
    Case Else
        ConvList = BaseStr(Index)
    End Select
End Property
Property Let ConvList(Index As Byte, Value As Byte)
    Select Case Index
    Case Is < 0, Is > 64
        'none
    Case Else
        SetBaseStr Index, Value
    End Select
End Property
Property Get Reserved() As Byte
    Reserved = BaseStr(64)
End Property
Property Let Reserved(Value As Byte)
    BaseStr(64) = Value
End Property

Private Sub Encrypt(ip1 As Byte, ip2 As Byte, ip3 As Byte, op1 As Byte, op2 As Byte, op3 As Byte, op4 As Byte)
    op1 = (ip1 And Bit72) \ 4
    op2 = ((ip1 And Bit10) * 16) Or ((ip2 And Bit74) \ 16)
    op3 = ((ip2 And Bit30) * 4) Or ((ip3 And Bit76) \ 64)
    op4 = ip3 And Bit50
End Sub
Private Sub Decrypt(ip1 As Byte, ip2 As Byte, ip3 As Byte, ip4 As Byte, op1 As Byte, op2 As Byte, op3 As Byte)
    op1 = (ip1 * 4) Or ((ip2 And Bit54) \ 16)
    op2 = ((ip2 And Bit30) * 16) Or ((ip3 And Bit52) \ 4)
    op3 = ((ip3 And Bit10) * 64) Or ip4
End Sub

Function DataEncode(ip() As Byte) As String
    On Error GoTo errED
    
    Dim ipS As Long, ipL As Long, ipU As Long
    ipL = LBound(ip): ipU = UBound(ip): ipS = ipU - ipL + 1
    
    RaiseEvent Init(ipS)
    
    Dim opS As Long, ipM As Byte, opR As Byte
    ipM = ipS Mod 3
    opR = (3 - ipM) Mod 3
    opS = ((ipS \ 3) + Abs(ipM <> 0)) * 4
    
    Dim op() As Byte
    ReDim op(0 To opS - 1)
    
    Dim B() As Byte
    ReDim B(0 To 3)
    
    Dim i&, j&, k&
    For i = ipL To ipU - 2 Step 3
        Encrypt ip(i), ip(i + 1), ip(i + 2), op(j), op(j + 1), op(j + 2), op(j + 3)
        j = j + 4
        If GetInputState Then
            RaiseEvent Progress(i)
            DoEvents
        End If
    Next i
    If ipM Then
        ReDim B(1 To 3)
        For k = 1 To ipM
            B(k) = ip(i)
            i = i + 1
        Next k
        Encrypt B(1), B(2), B(3), op(j), op(j + 1), op(j + 2), op(j + 3)
    End If
'    DataEncode = BinToTxt(MidB(op, 1, opS - opR)) & String(opR, BaseStr(64))
    DataEncode = BinToTxt(op) & String(opR, BaseStr(64))
    
    RaiseEvent Over
    Exit Function
errED:
    RaiseEvent Error(Err.Number, Err.Description)
    RaiseEvent Over
End Function
Private Function BinToTxt(ip() As Byte) As String
'On Error GoTo errB2T
    Dim i As Long, B() As Byte
    ReDim B(LBound(ip) To UBound(ip))
    For i = LBound(ip) To UBound(ip)
        B(i) = BaseStr(ip(i))
    Next i
    BinToTxt = StrConv(B, vbUnicode)
'errB2T:
End Function

Function DataDecode(ipStr As String) As Byte()
'    On Error GoTo errDD
    
    Dim ipS As Long, ip() As Byte
    ipS = Len(ipStr)
    
    RaiseEvent Init(ipS)
    
    TxtToBin ipStr, ip
    
    Dim op() As Byte
    ReDim op(0 To Int(ipS / 4) * 3 - 1)
    
    Dim opM As Byte, tP As Long, tS As String * 1
    tS = Chr(BaseStr(64))
    Do
        tP = InStr(tP + 1, ipStr, tS)
        If tP = 0 Then Exit Do
        opM = opM + 1
    Loop
    
    Dim i&, j&
    For i = LBound(ip) To UBound(ip) - opM Step 4
        Decrypt ip(i), ip(i + 1), ip(i + 2), ip(i + 3), op(j), op(j + 1), op(j + 2)
        j = j + 3
        If GetInputState Then
            RaiseEvent Progress(i)
            DoEvents
        End If
    Next i
    If opM Then
        DataDecode = MidB(op, 1, UBound(op) - LBound(op) + 1 - opM)
    Else
        DataDecode = op
    End If
    
    RaiseEvent Over
    Exit Function
errDD:
    RaiseEvent Error(Err.Number, Err.Description)
    RaiseEvent Over
End Function
Private Sub TxtToBin(ipText As String, opByte() As Byte)
    opByte = StrConv(ipText, vbFromUnicode)
    Dim i As Long
    For i = LBound(opByte) To UBound(opByte)
        opByte(i) = BaseRev(opByte(i))
    Next i
End Sub

Function TextEncode(ipText As String) As String
    On Error Resume Next
    If Len(ipText) Then TextEncode = DataEncode(StrConv(ipText, vbFromUnicode))
End Function
Function TextDecode(ipStr As String) As String
    On Error Resume Next
    If Len(ipStr) Then TextDecode = StrConv(DataDecode(ipStr), vbUnicode)
End Function
