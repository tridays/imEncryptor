Module mBase64
    Function BinToB64(ByRef Input As Byte()) As String
        Dim Output As String
        Try
            Output = Convert.ToBase64String(Input)
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            Output = Nothing
        End Try
        Return Output
    End Function
    Function B64ToBin(ByRef Input As String) As Byte()
        Return Convert.FromBase64String(Input)
    End Function

    Function StrToB64(ByRef Input As String) As String
        Return BinToB64(System.Text.Encoding.GetEncoding(-0).GetBytes(Input))
    End Function
    Function B64ToStr(ByRef Input As String) As String
        Dim Output As String
        Try
            Output = System.Text.Encoding.GetEncoding(-0).GetString(B64ToBin(Input))
        Catch ex As Exception
            MsgBox(ex.Message, 48)
            Output = Nothing
        End Try
        Return Output
    End Function
End Module
