Module mHex2Bin
    Function BinToHex(ByRef BinArray As Byte()) As String
        Dim i As Integer, s() As String
        ReDim s(BinArray.Length - 1)
        For i = 0 To BinArray.Length - 1
            s(i) = Hex(BinArray(i)).PadLeft(2, "0").ToLower
        Next
        Return Join(s, Nothing)
    End Function
    Function HexToBin(ByRef Text As String) As Byte()
        Dim i As Integer, j As Integer, B() As Byte, u As Integer
        u = Len(Text) \ 2 + Len(Text) Mod 2 - 1 '有余数其实说明输入格式错误
        ReDim B(u) : j = 0
        For i = 1 To Len(Text) Step 2
            B(j) = CByte("&H" + Mid(Text, i, 2))
            j += 1
        Next
        Return B
    End Function
End Module
