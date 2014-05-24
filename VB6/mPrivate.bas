Attribute VB_Name = "mPrivate"
Option Explicit

Private Declare Sub CpyMem4 Lib "msvbvm60.dll" Alias "GetMem4" (ByVal SrcPtr As Any, ByVal DstPtr As Any)

Private Declare Function GetInputState Lib "user32" () As Long

Dim Key As Long

Sub GetKey(ByVal Password As String)
    Dim Hash As New cHash
    Dim B() As Byte
    Dim i As Long, t As Long
    Hash.Algorithm = MD5
    If Len(Password) = 0 Then Password = "Password"
    B = HexToBin(Hash.TextHash(Password))
    Key = 0
    For i = 0 To 15 Step 4
        CpyMem4 VarPtr(B(i)), VarPtr(t)
        Key = Key Xor t
    Next i
    If Key = 0 Then Key = 123456789  '确保有种子
End Sub

Sub Encrypt(B() As Byte)
    Dim i As Long
    Rnd -Abs(Key)
    For i = 0 To UBound(B)
        B(i) = B(i) Xor Int(Rnd * 256)
        If GetInputState Then
            FrmMain.Jdt.Value = i
            DoEvents
        End If
    Next i
End Sub

'Sub Encrypt(B() As Byte)
'    Dim i As Long
'    Dim K(0 To 255) As Byte
'    Rnd -Abs(Key)
'    For i = 0 To UBound(K)
'        K(i) = CByte(Int(Rnd * 256)) '生成密钥盒
'    Next i
'    For i = 0 To UBound(B)
'        B(i) = B(i) Xor K(i Mod 256)
'        If GetInputState Then
'            FrmMain.Jdt.Value = i
'            DoEvents
'        End If
'    Next i
'End Sub

