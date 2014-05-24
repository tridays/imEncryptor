Module mStr2Bin
    Function StrToBin(ByRef s) As Byte()
        Return System.Text.Encoding.GetEncoding(-0).GetBytes(s)
    End Function
    Function BinToStr(ByRef BinArray() As Byte) As String
        Return System.Text.Encoding.GetEncoding(-0).GetString(BinArray)
    End Function
End Module
