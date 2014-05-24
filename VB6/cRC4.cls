VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'======================================================================================
'Name_ENG   RC4 Encryption
'Name_CHS   RC4 加密
'--------------------------------------------------------------------------------------
'Tips       源代码来自互联网
'======================================================================================

Private Declare Function GetInputState Lib "user32" () As Long

Dim S() As Byte

Dim mCount As Long
Dim mKey() As Byte

Private Sub Class_Initialize()
    mCount = 255
    mKey = StrConv("Password", vbFromUnicode)
End Sub

Property Get Count() As Long
    Count = mCount
End Property
Property Let Count(lngValue As Long)
    mCount = lngValue
End Property

Property Get Key() As Byte()
    Key = mKey
End Property
Property Let Key(bytValue() As Byte)
    mKey = bytValue
End Property

Sub RC4(Buffer() As Byte)
    Dim i&, j&, k&, x&, y&
    
    ReDim S(0 To mCount)
    For i = 0 To UBound(S)
        S(i) = i And &HFF
    Next i
    j = LBound(mKey)
    For i = 0 To UBound(S)
        k = (k + S(i) + mKey(j)) Mod mCount
        SwapByte S(i), S(k)
        j = j + 1
        If j = UBound(mKey) Then j = LBound(mKey)
    Next i
    
    For i = LBound(Buffer) To UBound(Buffer)
        x = (x + 1) Mod mCount
        y = (y + S(x)) Mod mCount
        SwapByte S(x), S(y)
        j = (CLng(S(x)) + S(y)) Mod mCount
        Buffer(i) = Buffer(i) Xor S(j)
    Next i
End Sub
Private Sub SwapByte(ByRef ip1 As Byte, ByRef ip2 As Byte)
    ip1 = ip1 Xor ip2
    ip2 = ip1 Xor ip2
    ip1 = ip1 Xor ip2
End Sub

Sub FileCrypt(Src As String, Dst As String)
    Dim FreeF As Integer, B() As Byte
    FreeF = FreeFile
    Open Src For Binary Access Read As #FreeF
    ReDim B(0 To LOF(FreeF) - 1)
    Get #FreeF, , B
    Close #FreeF
    Open Dst For Output As #FreeF
    Close #FreeF
    Open Dst For Binary Access Write As #FreeF
    Put #FreeF, , DataCrypt(B)
    Close #FreeF
End Sub

Function TextEncrypt(Src As String) As String
    TextEncrypt = BinToHex(DataCrypt(StrConv(Src, vbFromUnicode)))
End Function
Function TextDecrypt(Src As String) As String
    TextDecrypt = StrConv(DataCrypt(HexToBin(Src)), vbUnicode)
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

Function DataCrypt(Src() As Byte) As Byte()
    Dim Dst() As Byte
    Dst = Src: RC4 Dst
    DataCrypt = Dst
End Function
