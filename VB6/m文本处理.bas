Attribute VB_Name = "m文本处理"
Option Explicit

'======================================================================================
'Name_ENG   Text Processing Module
'Name_CHS   文本处理模块
'Version    1.3
'Author     Tridays
'Studio     Aimplus Studio
'Date       12-08-25
'--------------------------------------------------------------------------------------
'感谢使用，请保留版权信息
'--------------------------------------------------------------------------------------
'Update     v1.3 12-08-25
'Tips       新增 Bin <-> Hex <-> Txt 功能（移植自类模块）
'
'Update     v1.2 12-01-14
'Tips       基本功能设计
'======================================================================================

'使用说明
'--------------------------------------------------------------------------------------
'函数列表
'----------------------------------------------------------------
'LeftE(源字符串,
'      匹配字符串,
'      匹配起点,
'      匹配方式,
'      是否从左向右匹配【默认是】,
'      查找成功时是否一并返回匹配字符串【默认否】,
'      失败时是否返回匹配字符串，将位于最右端【默认否】,
'      失败时是否返回源字符串【默认否】
'      )
'      成功时返回匹配字符串之前的字符串内容（其他需求详见参数）

'LeftL(源字符串,
'      从右往左删除字符数
'      )
'      成功时返回去掉最右端n个字符后的字符串

'RightE(源字符串,
'       匹配字符串,
'       匹配起点,
'       匹配方式,
'       是否从左向右匹配【默认否】,
'       查找成功时是否一并返回匹配字符串【默认否】,
'       失败时是否返回匹配字符串，将位于最左端【默认否】,
'       失败时是否返回源字符串【默认否】
'       )
'       成功时输出匹配字符串之后的字符串内容（其他需求详见参数）

'RightL(源字符串,
'       从左往右删除字符数
'       )
'       成功时返回去掉最左端n个字符后的字符串

'MidE(源字符串,
'     左端匹配字符串,
'     右端匹配字符串,
'     匹配起点,
'     匹配方式,
'     是否从左向右匹配，仅对左端匹配字符串有效【默认是】,
'     查找成功时是否一并返回左端匹配字符串【默认否】,
'     查找成功时是否一并返回右端匹配字符串【默认否】,
'     失败时是否返回左端匹配字符串，将位于最左端【默认否】,
'     失败时是否返回右端匹配字符串，将位于最右端【默认否】,
'     失败时是否返回源字符串【默认否】
'     )
'     成功时输出匹配字符串之间的字符串内容（其他需求详见参数）

'MidL(源字符串,
'     左起绝对位置1,
'     左起绝对位置2
'     )
'     成功时返回位置1~位置2间的字符串
'--------------------------------------------------------------------------------------

'获取输入状态，非零时执行DoEvents防死
Private Declare Function GetInputState Lib "user32" () As Long

Function LeftE(strInput As String, _
               strMatch As String, _
      Optional lngStart As Long, _
      Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
      Optional FromLeft_ToRight As Boolean = True, _
      Optional Exist_Then_Add_strMatch As Boolean, _
      Optional Null_Then_Add_strMatch As Boolean, _
      Optional Null_Then_Reserved As Boolean _
               ) As String
    
    Dim Ret As String '存放结果
    If Len(strInput) = 0 Then Exit Function '使用Len()函数比字符串对比效率高
    If Len(strMatch) = 0 Then GoTo retNull
    Select Case lngStart                    '使用Select是因为能一值对多值
    Case Is < 0, Is > Len(strInput)
        GoTo retNull
    Case 0
        If FromLeft_ToRight Then '向左或向右查找字符串的函数初始值不同
            lngStart = 1
        Else
            lngStart = -1
        End If
    End Select
    
    Dim Pos As Long '存放截取位置
    If FromLeft_ToRight Then
        Pos = InStr(lngStart, strInput, strMatch, CompareMethod)     '从字符串第lngStart位往右匹配，返回从字符串首位起的位置
    Else
        Pos = InStrRev&(strInput, strMatch, lngStart, CompareMethod) '从字符串第lngStart位往左匹配，返回从字符串首位起的位置
    End If
    If Pos = 0 Then GoTo retNull                                '未找到
    If Exist_Then_Add_strMatch Then Pos = Pos + Len(strMatch)   '是否需加匹配字符串
    Ret = Left$(strInput, Pos - 1)
retNull:
    If Len(Ret) = 0 Then
        If Null_Then_Reserved Then Ret = strInput           '是否需保留源字符串
        If Null_Then_Add_strMatch Then Ret = Ret & strMatch '是否需保留匹配字符串
    End If
    LeftE = Ret
End Function
Function LeftL(strInput As String, SkipLen As Long) As String
    On Error Resume Next
    LeftL = Left$(strInput, Len(strInput) - SkipLen)
End Function

Function RightE(strInput As String, _
                strMatch As String, _
       Optional lngStart As Long, _
       Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
       Optional FromLeft_ToRight As Boolean = False, _
       Optional Exist_Then_Add_strMatch As Boolean, _
       Optional Null_Then_Add_strMatch As Boolean, _
       Optional Null_Then_Reserved As Boolean _
                ) As String
    
    Dim Ret As String '存放结果
    If Len(strInput) = 0 Then Exit Function '使用Len()函数比字符串对比效率高
    If Len(strMatch) = 0 Then GoTo retNull
    Select Case lngStart                    '使用Select是因为能一值对多值
    Case Is < 0, Is > Len(strInput)
        GoTo retNull
    Case 0
        If FromLeft_ToRight Then '向左或向右查找字符串的函数初始值不同
            lngStart = 1
        Else
            lngStart = -1
        End If
    End Select
    
    Dim Pos As Long '存放截取位置
    If FromLeft_ToRight Then
        Pos = InStr(lngStart, strInput, strMatch, CompareMethod)     '从字符串第lngStart位往右匹配，返回从字符串首位起的位置
    Else
        Pos = InStrRev&(strInput, strMatch, lngStart, CompareMethod) '从字符串第lngStart位往左匹配，返回从字符串首位起的位置
    End If
    If Pos = 0 Then GoTo retNull                                        '未找到
    If Exist_Then_Add_strMatch = False Then Pos = Pos + Len(strMatch)   '是否需去掉匹配字符串
    Ret = Right$(strInput, Len(strInput) - Pos + 1)
retNull:
    If Len(Ret) = 0 Then
        If Null_Then_Reserved Then Ret = strInput           '是否需保留源字符串
        If Null_Then_Add_strMatch Then Ret = strMatch & Ret '是否需保留匹配字符串
    End If
    RightE = Ret
End Function
Function RightL(strInput As String, SkipLen As Long) As String
    '在这里感觉上使用Mid函数更高效，注释掉的部分为最初敲定的本函数代码
    On Error Resume Next
'    RightL = Right$(strInput, Len(strInput) - SkipLen)
    RightL = Mid$(strInput, SkipLen + 1)
End Function

Function MidE(strInput As String, _
              strMatch1 As String, _
              strMatch2 As String, _
     Optional lngStart As Long, _
     Optional CompareMethod As VbCompareMethod = vbBinaryCompare, _
     Optional FromLeft_ToRight As Boolean = True, _
     Optional Exist_Then_Add_strMatch1 As Boolean, _
     Optional Exist_Then_Add_strMatch2 As Boolean, _
     Optional Null_Then_Add_strMatch1 As Boolean, _
     Optional Null_Then_Add_strMatch2 As Boolean, _
     Optional Null_Then_Reserved As Boolean _
              ) As String

    Dim Ret As String '存放结果
    If Len(strInput) = 0 Then Exit Function '使用Len()函数比字符串对比效率高
    If Len(strMatch1) = 0 Then
        Ret = LeftE(strInput, strMatch2, lngStart, CompareMethod, FromLeft_ToRight, Exist_Then_Add_strMatch2, Null_Then_Add_strMatch2, Null_Then_Reserved)
        GoTo retNull
    ElseIf Len(strMatch2) = 0 Then
        Ret = RightE(strInput, strMatch1, lngStart, CompareMethod, FromLeft_ToRight, Exist_Then_Add_strMatch2, Null_Then_Add_strMatch2, Null_Then_Reserved)
        GoTo retNull
    End If
    Select Case lngStart                    '使用Select是因为能一值对多值
    Case Is < 0, Is > Len(strInput)
        GoTo retNull
    Case 0
        If FromLeft_ToRight Then '向左或向右查找字符串的函数初始值不同
            lngStart = 1
        Else
            lngStart = -1
        End If
    End Select
    
    Dim Pos1 As Long, Pos2 As Long '存放截取位置
    If FromLeft_ToRight Then
        Pos1 = InStr(lngStart, strInput, strMatch1, CompareMethod)       '从字符串第lngStart位往右匹配，返回从字符串首位起的位置
    Else
        Pos1 = InStrRev&(strInput, strMatch1, lngStart, CompareMethod)   '从字符串第lngStart位往左匹配，返回从字符串首位起的位置
    End If
    If Pos1 = 0 Then GoTo retNull                               '未找到
    If Exist_Then_Add_strMatch1 = False Then Pos1 = Pos1 + Len(strMatch1)   '是否需加左端匹配字符串（不需要跳过）
    Pos2 = InStr(Pos1, strInput, strMatch2, CompareMethod)
    If Pos2 = 0 Then GoTo retNull
    If Exist_Then_Add_strMatch2 Then Pos2 = Pos2 + Len(strMatch2)           '是否需加右端匹配字符串（需要就补上）
    Ret = Mid$(strInput, Pos1, Pos2 - Pos1)
retNull:
    If Len(Ret) = 0 Then
        If Null_Then_Reserved Then Ret = strInput               '是否需保留源字符串
        If Null_Then_Add_strMatch1 Then Ret = strMatch1 & Ret   '是否需保留左端匹配字符串
        If Null_Then_Add_strMatch2 Then Ret = Ret & strMatch2   '是否需保留右端匹配字符串
    End If
    MidE = Ret
End Function
Function MidL(strInput As String, Pos1 As Long, Pos2 As Long) As String
    On Error Resume Next
    If Pos2 >= Pos1 Then MidL = Mid$(strInput, Pos1, Pos2 - Pos1 + 1)
End Function

'计算匹配字符串在原字符串中出现频数
Function TimesInStr(strInput As String, strMatch As String) As Long
    Dim Pos As Long
    Do
        Pos = InStr(Pos + 1, strInput, strMatch)
        TimesInStr = TimesInStr + 1
    Loop While Pos
    TimesInStr = TimesInStr - 1
End Function

'一一对应地替换源字符串
Function RepEx(ByVal strInput As String, FindList As String, RepList As String) As String
    Dim i&
    For i = 1 To Len(RepList)
        strInput = Replace$(strInput, Mid$(FindList, i, 1), Mid$(RepList, i, 1))
    Next i
    RepEx = strInput
End Function

'读写文本文档，支持代码页转换（例如GB2312转换为UTF-8）
Function LoadTxtEx(Path As String, Optional Charset As String = "GB2312") As String
'    Dim Aso As New ADODB.Stream
    Dim Aso As Object
    Set Aso = CreateObject("ADODB.Stream")
    With Aso
        .Type = 2 'adTypeText
        .Mode = 3 'adModeReadWrite
        .Charset = Charset
        .Open
        .LoadFromFile Path
        LoadTxtEx = .ReadText
        .Close
    End With
    Set Aso = Nothing
End Function
Sub SaveTxtEx(Text As String, Path As String, Optional Charset As String = "GB2312")
'    Dim Aso As New ADODB.Stream
    Dim Aso As Object
    Set Aso = CreateObject("ADODB.Stream")
    With Aso
        .Type = 2 'adTypeText
        .Mode = 3 'adModeReadWrite
        .Charset = Charset
        .Open
        .WriteText Text
        .SaveToFile Path, 2 'adSaveCreateOverWrite
        .Flush
        .Close
    End With
    Set Aso = Nothing
End Sub

'------------------------------------------------------------
'二进制数组 ←→ 十六进制文本 ←→ 正常文本
'------------------------------------------------------------

Function BinToHex(B() As Byte) As String
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
Function HexToBin(H As String, Optional lngLBound As Long) As Byte()
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

Function TxtToHex(t As String) As String
    If Len(t) Then TxtToHex = BinToHex(StrConv(t, vbFromUnicode))
End Function
Function HexToTxt(H As String) As String
    If Len(H) Then
        Dim t$
        t = HexToBin(H)
        If Len(t) Then HexToTxt = StrConv(t, vbUnicode)
    End If
End Function
