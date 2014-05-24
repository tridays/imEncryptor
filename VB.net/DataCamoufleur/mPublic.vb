Module mPublic
    Public Declare Function GetInputState Lib "user32" () As Integer

    Public mKey As Integer

    Function inIDE() As Boolean
        Dim bol As Boolean = False
        Debug.Assert(SetTrue(bol))
        Return bol
    End Function
    Private Function SetTrue(ByRef bol As Boolean) As Boolean
        bol = True
        Return True
    End Function

    Sub GetKey(ByVal Password As String)
        If Len(Password) = 0 Then Password = "Password"
        Dim i As Integer, B() As Byte, si As Integer
        B = HexToBin(MD5(Password, 32))
        mKey = 0
        For i = 0 To 15 Step 4
            si = System.BitConverter.ToInt32(B, i)
            mKey = mKey Xor si
        Next
        mKey = Math.Abs(mKey)
        If mKey = 0 Then mKey = 1234567890 '确保有种子
    End Sub
    Private Function MD5(ByVal Text As String, Optional ByVal Code As Int16 = 32) As String
        Dim B As Byte() = (New System.Text.ASCIIEncoding).GetBytes(Text)
        Dim v As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(B)
        Dim rtn As String = Nothing
        Dim i As Integer
        Select Case Code
            Case 16   '16位字符   
                For i = 4 To 11
                    rtn &= Hex(v(i)).PadLeft(2, "0").ToLower
                Next
            Case Else '32位字符（默认）
                For i = 0 To 15
                    rtn &= Hex(v(i)).PadLeft(2, "0").ToLower
                Next
        End Select
        Return rtn
    End Function
End Module
