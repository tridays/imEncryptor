Attribute VB_Name = "m�ı�����"
Option Explicit

'======================================================================================
'Name_ENG   Text Processing Module
'Name_CHS   �ı�����ģ��
'Version    1.3
'Author     Tridays
'Studio     Aimplus Studio
'Date       12-08-25
'--------------------------------------------------------------------------------------
'��лʹ�ã��뱣����Ȩ��Ϣ
'--------------------------------------------------------------------------------------
'Update     v1.3 12-08-25
'Tips       ���� Bin <-> Hex <-> Txt ���ܣ���ֲ����ģ�飩
'
'Update     v1.2 12-01-14
'Tips       �����������
'======================================================================================

'ʹ��˵��
'--------------------------------------------------------------------------------------
'�����б�
'----------------------------------------------------------------
'LeftE(Դ�ַ���,
'      ƥ���ַ���,
'      ƥ�����,
'      ƥ�䷽ʽ,
'      �Ƿ��������ƥ�䡾Ĭ���ǡ�,
'      ���ҳɹ�ʱ�Ƿ�һ������ƥ���ַ�����Ĭ�Ϸ�,
'      ʧ��ʱ�Ƿ񷵻�ƥ���ַ�������λ�����Ҷˡ�Ĭ�Ϸ�,
'      ʧ��ʱ�Ƿ񷵻�Դ�ַ�����Ĭ�Ϸ�
'      )
'      �ɹ�ʱ����ƥ���ַ���֮ǰ���ַ������ݣ������������������

'LeftL(Դ�ַ���,
'      ��������ɾ���ַ���
'      )
'      �ɹ�ʱ����ȥ�����Ҷ�n���ַ�����ַ���

'RightE(Դ�ַ���,
'       ƥ���ַ���,
'       ƥ�����,
'       ƥ�䷽ʽ,
'       �Ƿ��������ƥ�䡾Ĭ�Ϸ�,
'       ���ҳɹ�ʱ�Ƿ�һ������ƥ���ַ�����Ĭ�Ϸ�,
'       ʧ��ʱ�Ƿ񷵻�ƥ���ַ�������λ������ˡ�Ĭ�Ϸ�,
'       ʧ��ʱ�Ƿ񷵻�Դ�ַ�����Ĭ�Ϸ�
'       )
'       �ɹ�ʱ���ƥ���ַ���֮����ַ������ݣ������������������

'RightL(Դ�ַ���,
'       ��������ɾ���ַ���
'       )
'       �ɹ�ʱ����ȥ�������n���ַ�����ַ���

'MidE(Դ�ַ���,
'     ���ƥ���ַ���,
'     �Ҷ�ƥ���ַ���,
'     ƥ�����,
'     ƥ�䷽ʽ,
'     �Ƿ��������ƥ�䣬�������ƥ���ַ�����Ч��Ĭ���ǡ�,
'     ���ҳɹ�ʱ�Ƿ�һ���������ƥ���ַ�����Ĭ�Ϸ�,
'     ���ҳɹ�ʱ�Ƿ�һ�������Ҷ�ƥ���ַ�����Ĭ�Ϸ�,
'     ʧ��ʱ�Ƿ񷵻����ƥ���ַ�������λ������ˡ�Ĭ�Ϸ�,
'     ʧ��ʱ�Ƿ񷵻��Ҷ�ƥ���ַ�������λ�����Ҷˡ�Ĭ�Ϸ�,
'     ʧ��ʱ�Ƿ񷵻�Դ�ַ�����Ĭ�Ϸ�
'     )
'     �ɹ�ʱ���ƥ���ַ���֮����ַ������ݣ������������������

'MidL(Դ�ַ���,
'     �������λ��1,
'     �������λ��2
'     )
'     �ɹ�ʱ����λ��1~λ��2����ַ���
'--------------------------------------------------------------------------------------

'��ȡ����״̬������ʱִ��DoEvents����
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
    
    Dim Ret As String '��Ž��
    If Len(strInput) = 0 Then Exit Function 'ʹ��Len()�������ַ����Ա�Ч�ʸ�
    If Len(strMatch) = 0 Then GoTo retNull
    Select Case lngStart                    'ʹ��Select����Ϊ��һֵ�Զ�ֵ
    Case Is < 0, Is > Len(strInput)
        GoTo retNull
    Case 0
        If FromLeft_ToRight Then '��������Ҳ����ַ����ĺ�����ʼֵ��ͬ
            lngStart = 1
        Else
            lngStart = -1
        End If
    End Select
    
    Dim Pos As Long '��Ž�ȡλ��
    If FromLeft_ToRight Then
        Pos = InStr(lngStart, strInput, strMatch, CompareMethod)     '���ַ�����lngStartλ����ƥ�䣬���ش��ַ�����λ���λ��
    Else
        Pos = InStrRev&(strInput, strMatch, lngStart, CompareMethod) '���ַ�����lngStartλ����ƥ�䣬���ش��ַ�����λ���λ��
    End If
    If Pos = 0 Then GoTo retNull                                'δ�ҵ�
    If Exist_Then_Add_strMatch Then Pos = Pos + Len(strMatch)   '�Ƿ����ƥ���ַ���
    Ret = Left$(strInput, Pos - 1)
retNull:
    If Len(Ret) = 0 Then
        If Null_Then_Reserved Then Ret = strInput           '�Ƿ��豣��Դ�ַ���
        If Null_Then_Add_strMatch Then Ret = Ret & strMatch '�Ƿ��豣��ƥ���ַ���
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
    
    Dim Ret As String '��Ž��
    If Len(strInput) = 0 Then Exit Function 'ʹ��Len()�������ַ����Ա�Ч�ʸ�
    If Len(strMatch) = 0 Then GoTo retNull
    Select Case lngStart                    'ʹ��Select����Ϊ��һֵ�Զ�ֵ
    Case Is < 0, Is > Len(strInput)
        GoTo retNull
    Case 0
        If FromLeft_ToRight Then '��������Ҳ����ַ����ĺ�����ʼֵ��ͬ
            lngStart = 1
        Else
            lngStart = -1
        End If
    End Select
    
    Dim Pos As Long '��Ž�ȡλ��
    If FromLeft_ToRight Then
        Pos = InStr(lngStart, strInput, strMatch, CompareMethod)     '���ַ�����lngStartλ����ƥ�䣬���ش��ַ�����λ���λ��
    Else
        Pos = InStrRev&(strInput, strMatch, lngStart, CompareMethod) '���ַ�����lngStartλ����ƥ�䣬���ش��ַ�����λ���λ��
    End If
    If Pos = 0 Then GoTo retNull                                        'δ�ҵ�
    If Exist_Then_Add_strMatch = False Then Pos = Pos + Len(strMatch)   '�Ƿ���ȥ��ƥ���ַ���
    Ret = Right$(strInput, Len(strInput) - Pos + 1)
retNull:
    If Len(Ret) = 0 Then
        If Null_Then_Reserved Then Ret = strInput           '�Ƿ��豣��Դ�ַ���
        If Null_Then_Add_strMatch Then Ret = strMatch & Ret '�Ƿ��豣��ƥ���ַ���
    End If
    RightE = Ret
End Function
Function RightL(strInput As String, SkipLen As Long) As String
    '������о���ʹ��Mid��������Ч��ע�͵��Ĳ���Ϊ����ö��ı���������
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

    Dim Ret As String '��Ž��
    If Len(strInput) = 0 Then Exit Function 'ʹ��Len()�������ַ����Ա�Ч�ʸ�
    If Len(strMatch1) = 0 Then
        Ret = LeftE(strInput, strMatch2, lngStart, CompareMethod, FromLeft_ToRight, Exist_Then_Add_strMatch2, Null_Then_Add_strMatch2, Null_Then_Reserved)
        GoTo retNull
    ElseIf Len(strMatch2) = 0 Then
        Ret = RightE(strInput, strMatch1, lngStart, CompareMethod, FromLeft_ToRight, Exist_Then_Add_strMatch2, Null_Then_Add_strMatch2, Null_Then_Reserved)
        GoTo retNull
    End If
    Select Case lngStart                    'ʹ��Select����Ϊ��һֵ�Զ�ֵ
    Case Is < 0, Is > Len(strInput)
        GoTo retNull
    Case 0
        If FromLeft_ToRight Then '��������Ҳ����ַ����ĺ�����ʼֵ��ͬ
            lngStart = 1
        Else
            lngStart = -1
        End If
    End Select
    
    Dim Pos1 As Long, Pos2 As Long '��Ž�ȡλ��
    If FromLeft_ToRight Then
        Pos1 = InStr(lngStart, strInput, strMatch1, CompareMethod)       '���ַ�����lngStartλ����ƥ�䣬���ش��ַ�����λ���λ��
    Else
        Pos1 = InStrRev&(strInput, strMatch1, lngStart, CompareMethod)   '���ַ�����lngStartλ����ƥ�䣬���ش��ַ�����λ���λ��
    End If
    If Pos1 = 0 Then GoTo retNull                               'δ�ҵ�
    If Exist_Then_Add_strMatch1 = False Then Pos1 = Pos1 + Len(strMatch1)   '�Ƿ�������ƥ���ַ���������Ҫ������
    Pos2 = InStr(Pos1, strInput, strMatch2, CompareMethod)
    If Pos2 = 0 Then GoTo retNull
    If Exist_Then_Add_strMatch2 Then Pos2 = Pos2 + Len(strMatch2)           '�Ƿ�����Ҷ�ƥ���ַ�������Ҫ�Ͳ��ϣ�
    Ret = Mid$(strInput, Pos1, Pos2 - Pos1)
retNull:
    If Len(Ret) = 0 Then
        If Null_Then_Reserved Then Ret = strInput               '�Ƿ��豣��Դ�ַ���
        If Null_Then_Add_strMatch1 Then Ret = strMatch1 & Ret   '�Ƿ��豣�����ƥ���ַ���
        If Null_Then_Add_strMatch2 Then Ret = Ret & strMatch2   '�Ƿ��豣���Ҷ�ƥ���ַ���
    End If
    MidE = Ret
End Function
Function MidL(strInput As String, Pos1 As Long, Pos2 As Long) As String
    On Error Resume Next
    If Pos2 >= Pos1 Then MidL = Mid$(strInput, Pos1, Pos2 - Pos1 + 1)
End Function

'����ƥ���ַ�����ԭ�ַ����г���Ƶ��
Function TimesInStr(strInput As String, strMatch As String) As Long
    Dim Pos As Long
    Do
        Pos = InStr(Pos + 1, strInput, strMatch)
        TimesInStr = TimesInStr + 1
    Loop While Pos
    TimesInStr = TimesInStr - 1
End Function

'һһ��Ӧ���滻Դ�ַ���
Function RepEx(ByVal strInput As String, FindList As String, RepList As String) As String
    Dim i&
    For i = 1 To Len(RepList)
        strInput = Replace$(strInput, Mid$(FindList, i, 1), Mid$(RepList, i, 1))
    Next i
    RepEx = strInput
End Function

'��д�ı��ĵ���֧�ִ���ҳת��������GB2312ת��ΪUTF-8��
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
'���������� ���� ʮ�������ı� ���� �����ı�
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
