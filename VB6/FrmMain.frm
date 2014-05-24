VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "msComCtl32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "msComDlg32.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "imEncryptor ver."
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8295
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8295
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdA 
      Caption         =   "退出"
      Height          =   615
      Index           =   0
      Left            =   6600
      TabIndex        =   37
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "关于"
      Height          =   615
      Index           =   1
      Left            =   5160
      TabIndex        =   36
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Frame Fmf 
      Caption         =   "文件 File"
      Height          =   1815
      Left            =   120
      TabIndex        =   32
      Top             =   6000
      Width           =   4935
      Begin VB.ComboBox cbF 
         Height          =   300
         ItemData        =   "FrmMain.frx":6852
         Left            =   120
         List            =   "FrmMain.frx":6862
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton CmdF 
         Caption         =   "…"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox tF 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   960
         OLEDropMode     =   1  'Manual
         TabIndex        =   16
         Top             =   720
         Width           =   3435
      End
      Begin VB.CommandButton CmdF 
         Caption         =   "…"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton CmdF 
         Caption         =   "解密"
         Height          =   495
         Index           =   1
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton CmdF 
         Caption         =   "加密"
         Height          =   495
         Index           =   0
         Left            =   1920
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox tF 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   960
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   240
         Width           =   3435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "加密格式："
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "目标址"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   780
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "源文件"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   33
         Top             =   300
         Width           =   720
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Fmp 
      Caption         =   "密钥 Passowrd"
      ForeColor       =   &H00008000&
      Height          =   1095
      Left            =   5160
      TabIndex        =   31
      Top             =   6000
      Width           =   3015
      Begin VB.CheckBox cbKey 
         Caption         =   "使用迷惑性字符代替文本显示"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.TextBox tK 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.PictureBox PicWait 
      Height          =   1335
      Left            =   3000
      ScaleHeight     =   1275
      ScaleWidth      =   2235
      TabIndex        =   29
      Top             =   2520
      Visible         =   0   'False
      Width           =   2295
      Begin MSComctlLib.ProgressBar Jdt 
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "请稍候"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Fmt 
      Caption         =   "文本 Text"
      Height          =   5775
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton CmdP 
         Caption         =   "加密"
         Height          =   735
         Index           =   2
         Left            =   3240
         TabIndex        =   6
         Top             =   4920
         Width           =   735
      End
      Begin VB.PictureBox pC 
         AutoRedraw      =   -1  'True
         Height          =   4095
         Left            =   4080
         ScaleHeight     =   4035
         ScaleWidth      =   3795
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
         Begin VB.Label lblPic 
            BackStyle       =   0  'Transparent
            Caption         =   "×"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   3360
            TabIndex        =   39
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.ComboBox cbP 
         Height          =   300
         ItemData        =   "FrmMain.frx":68A6
         Left            =   1080
         List            =   "FrmMain.frx":68B3
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   5400
         Width           =   2055
      End
      Begin VB.ComboBox cbC 
         Height          =   300
         ItemData        =   "FrmMain.frx":68E7
         Left            =   5040
         List            =   "FrmMain.frx":68F4
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton CmdC 
         Caption         =   "导出"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   10
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdC 
         Caption         =   "导入"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   9
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdP 
         Caption         =   "→"
         Height          =   375
         Index           =   4
         Left            =   2400
         TabIndex        =   4
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdP 
         Caption         =   "导出"
         Height          =   375
         Index           =   1
         Left            =   840
         TabIndex        =   2
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdP 
         Caption         =   "导入"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdP 
         Caption         =   "清空"
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   3
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdC 
         Caption         =   "←"
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   8
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdC 
         Caption         =   "清空"
         Height          =   375
         Index           =   3
         Left            =   6360
         TabIndex        =   11
         Top             =   4920
         Width           =   735
      End
      Begin VB.CommandButton CmdC 
         Caption         =   "解密"
         Height          =   735
         Index           =   2
         Left            =   7200
         TabIndex        =   13
         Top             =   4920
         Width           =   735
      End
      Begin RichTextLib.RichTextBox tP 
         Height          =   4095
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7223
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"FrmMain.frx":6928
      End
      Begin RichTextLib.RichTextBox tC 
         Height          =   4095
         Left            =   4080
         TabIndex        =   7
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   7223
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   2
         TextRTF         =   $"FrmMain.frx":69C5
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输出格式："
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   5460
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密文格式："
         Height          =   180
         Left            =   4080
         TabIndex        =   26
         Top             =   5460
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "明文 PlainText"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密文 CipherText"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   4200
         TabIndex        =   24
         Top             =   360
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CpyMem4 Lib "msvbvm60.dll" Alias "GetMem4" (ByVal SrcPtr As Any, ByVal DstPtr As Any)

Private Declare Function GetInputState Lib "user32" () As Long

Private WithEvents B64 As cBase64
Attribute B64.VB_VarHelpID = -1
Dim BMP As cBitmap
Dim Fso As FileSystemObject

'-----------------------------------------------------------------------------------
Dim Key As Long

Private Sub GetKey(ByVal Password As String)
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

Private Sub Encrypt(B() As Byte)
    Dim i As Long
    Rnd -Abs(Key)
    For i = 0 To UBound(B)
        B(i) = B(i) Xor Int(Rnd * 256)
        If GetInputState Or (i Mod 102400 = 0) Then
            PicWait.Visible = True
            Jdt.Value = i
            DoEvents
        End If
    Next i
End Sub
'-----------------------------------------------------------------------------------

Private Sub B64_Error(Number As Integer, Description As String)
    MsgBox Number & ":" & Description, 16
End Sub
Private Sub B64_Init(Duration As Long)
    PicWait.Visible = True
    Jdt.Max = Duration
End Sub
Private Sub B64_Over()
    Jdt.Value = Jdt.Max
    PicWait.Visible = False
End Sub
Private Sub B64_Progress(Position As Long)
    Jdt.Value = Position
End Sub

Private Sub cbC_Click()
    If pC.Visible = True And (cbC.ListIndex <> 2) Then
        If MsgBox("这个操作将取消对图片的操作，确认吗？", 36) = vbYes Then
            pC.Visible = False
        Else
            cbC.ListIndex = 2
        End If
    End If
End Sub

Private Sub cbKey_Click()
    tK.PasswordChar = IIf(cbKey.Value = 1, "*", "")
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & App.Major & "." & App.Minor
    cbP.ListIndex = 0
    cbC.ListIndex = 0
    cbF.ListIndex = 0
    
    Set B64 = New cBase64
    Set BMP = New cBitmap
    Set Fso = New FileSystemObject
    
'    Dim strCMD As String
'    strCMD = Command
'    If strCMD <> "" Then
'        CorrectPath strCMD
'        tF(0).Text = strCMD: tF_LostFocus 0
'        tF(1).Text = strCMD: tF_LostFocus 1
'        Select Case MsgBox("请选择快捷操作：[是]编码[否]解码[取消]退出。", vbYesNoCancel Or vbQuestion)
'            Case vbYes: CmdF_Click 0
'            Case vbNo:  CmdF_Click 1
'            Case vbCancel 'none
'        End Select
'        End
'    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If PicWait.Visible = True Then
        If MsgBox("正在计算，要强制退出吗？", 36) = vbNo Then Cancel = 1
    End If
End Sub

Private Sub CmdA_Click(Index As Integer)
    Select Case Index
        Case 0 '退出
            Unload Me
        Case 1 '关于
            FrmAboutMe.Show
    End Select
End Sub

Private Sub CmdP_Click(Index As Integer) '明文操作
    On Error GoTo errCancel
    With CD
        .Filter = "文本文档(*.txt)|*.txt|所有文件(*.*)|*"
        .FilterIndex = 0
        Select Case Index
            Case 0 '导入
                .Flags = cdlOFNFileMustExist
                .FileName = ""
                .ShowOpen
                tP.LoadFile .FileName, rtfText
            Case 1 '导出
                .Flags = cdlOFNOverwritePrompt
                .ShowSave
                tP.SaveFile .FileName, rtfText
            Case 2 '加密
                If Len(tP.Text) = 0 Then Exit Sub
                If cbP.ListIndex = 2 Then
                    .Filter = "BMP位图(*.bmp)|*.bmp|所有文件(*.*)|*"
                    .Flags = cdlOFNOverwritePrompt
                    .FileName = "*.bmp"
                    .ShowSave
'                ElseIf Len(tC.Text) Then
'                    If MsgBox("即将替换掉当前“密文框”的文本，确认操作？", 36) = vbNo Then Exit Sub
                End If
                
                On Error GoTo 0
                GetKey tK.Text
                Dim B() As Byte
                B = StrConv(tP.Text, vbFromUnicode)
                With Jdt
                    .Value = 0
                    .Max = UBound(B) + 1
                End With
                PicWait.Visible = True
                
                Encrypt B
                
                Select Case cbP.ListIndex
                Case 0
                    pC.Visible = False
                    tC.Text = B64.DataEncode(B)
                Case 1
                    pC.Visible = False
                    tC.Text = BinToHex(B)
                Case 2
                    BMP.GetDataFromArray B
                    BMP.SaveBitmap CD.FileName
                    Set pC.Picture = LoadPicture(CD.FileName)
                    pC.Visible = True
                End Select
                
                With Jdt
                    .Value = .Max
                End With
                PicWait.Visible = False
                cbC.ListIndex = cbP.ListIndex
                MsgBox "加密完成，请查看“密文框”！", 64
            Case 3 '清空
                If Len(tP.Text) Then If MsgBox("确认清空“明文框”？", 36) = vbYes Then tP.Text = ""
            Case 4 '转移
                If Len(tP.Text) = 0 Then Exit Sub
                If Len(tC.Text) Then If MsgBox("即将把文本复制到“密文框”，确认操作？", 36) = vbNo Then Exit Sub
                tC.Text = tP.Text: tP.Text = ""
        End Select
    End With
errCancel:
End Sub
Private Sub CmdC_Click(Index As Integer) '密文操作
    On Error GoTo errCancel
    With CD
        .Filter = "文本文档(*.txt)|*.txt|BMP位图(*.bmp)|*.bmp|所有文件(*.*)|*"
        .FilterIndex = IIf(cbC.ListIndex = 2, 2, 1)
        Select Case Index
            Case 0 '导入
                .Flags = cdlOFNFileMustExist
                .FileName = ""
                .ShowOpen
                If .FilterIndex = 2 Then cbC.ListIndex = 2
                If cbC.ListIndex = 2 Then
                    pC.Visible = True
                    Set pC.Picture = LoadPicture(CD.FileName)
                    BMP.LoadBitmap CD.FileName
                Else
                    pC.Visible = False
                    tC.LoadFile .FileName, rtfText
                End If
            Case 1 '导出
                .Flags = cdlOFNOverwritePrompt
                .ShowSave
                tC.SaveFile .FileName, rtfText
            Case 2 '解密
                If cbC.ListIndex = 2 And pC.Visible = False Then
                    CD.FileName = ""
                    CmdC_Click 0
                    If Len(CD.FileName) Then CmdC_Click 2
                    Exit Sub
                End If
                If cbC.ListIndex <> 2 And Len(tC.Text) = 0 Then Exit Sub
'                If Len(tP.Text) Then If MsgBox("即将替换掉当前“明文框”的文本，确认操作？", 36) = vbNo Then Exit Sub
                
                On Error GoTo 0
                GetKey tK.Text
                Dim B() As Byte
                Select Case cbC.ListIndex
                Case 0 'BASE64编码文本
                    B = B64.DataDecode(tC.Text)
                Case 1 '十六进制字符串
                    B = HexToBin(tC.Text)
                Case 2 'ＢＭＰ格式图片
                    If pC.Visible = False Then '如果没有打开图片
                        CmdC_Click 0
                    End If
                    BMP.PutDataToArray B
                End Select
                
                '初始化界面
                With Jdt
                    .Value = 0
                    .Max = UBound(B) + 1
                End With
                PicWait.Visible = True
                
                Encrypt B '加密/解密
                
                '界面收尾
                tP.Text = StrConv(B, vbUnicode)
                With Jdt
                    .Value = .Max
                End With
                PicWait.Visible = False
                
                MsgBox "解密完成，请查看“明文框”！", 64
            Case 3 '清空
                pC.Visible = False
                If Len(tC.Text) Then If MsgBox("确认清空“密文框”？", 36) = vbYes Then tC.Text = ""
            Case 4 '转移
                If Len(tC.Text) = 0 Then Exit Sub
                If Len(tP.Text) Then If MsgBox("即将把文本复制到“明文框”，确认操作？", 36) = vbNo Then Exit Sub
                tP.Text = tC.Text: tC.Text = ""
        End Select
    End With
errCancel:
End Sub

Private Sub CmdF_Click(Index As Integer)
    On Error GoTo errCancel
    With CD
        .Filter = "所有文件(*.*)|*"
        .FilterIndex = 0
        Select Case Index
            Case 0, 1 '加、解密
                Dim i As Long
                For i = tF.LBound To tF.ubound
                    If tF(i).ForeColor <> RGB(0, 0, 255) Then
                        MsgBox "请核实路径！", 48
                        GoTo errCancel
                    End If
                Next i
                
                On Error GoTo 0
                GetKey tK.Text
                Dim B() As Byte
                Dim BM2 As New cBitmap
                With Jdt
                    .Value = 0
                    .Max = FileLen(tF(0).Text)
                End With
                PicWait.Visible = True
                
                If Index = 0 Then   '加密
                    B = LoadBin(tF(0).Text)
                    Encrypt B
                    Select Case cbF.ListIndex
                    Case 0
                        SaveBin B, tF(1).Text
                    Case 1
                        SaveTxt B64.DataEncode(B), tF(1).Text
                    Case 2
                        SaveTxt BinToHex(B), tF(1).Text
                    Case 3
                        BM2.GetDataFromArray B
                        BM2.SaveBitmap tF(1).Text
                    End Select
                Else                '解密
                    Select Case cbF.ListIndex
                    Case 0
                        B = LoadBin(tF(0).Text)
                    Case 1
                        B = B64.DataDecode(LoadTxt(tF(0).Text))
                    Case 2
                        B = HexToBin(LoadTxt(tF(0).Text))
                    Case 3
                        BM2.LoadBitmap tF(0).Text
                        BM2.PutDataToArray B
                    End Select
                    Encrypt B
                    SaveBin B, tF(1).Text
                End If
                PicWait.Visible = False
                With Jdt
                    .Value = .Max
                End With
                MsgBox "操作完成！", 64
            Case 2 '源
                .Flags = cdlOFNFileMustExist
                .FileName = ""
                .ShowOpen
                If CheckFile(.FileName) = 0 Then
                    tF(0).Text = .FileName
                    tF(0).ForeColor = RGB(0, 0, 255)
                    
                    tF(0).SelStart = Len(tF(0).Text)
                    If Len(tF(1).Text) = 0 Then
                        If cbF.ListIndex = 2 Then
                            tF(1).Text = tF(0).Text & ".bmp"
                        Else
                            tF(1).Text = tF(0).Text
                        End If
                        tF(1).SelStart = Len(tF(1).Text)
                    End If
                End If
            Case 3 '目标
                .Flags = cdlOFNOverwritePrompt
                .ShowSave
                If CheckFile(.FileName) <> 2 Then
                    tF(1).Text = .FileName
                    tF(1).ForeColor = RGB(0, 0, 255)
                End If
        End Select
    End With
errCancel:
End Sub

Private Sub lblPic_Click()
    If MsgBox("即将取消对当前图片的操作，返回密文框界面。" & vbCrLf & vbCrLf & "确认操作？", 36) = vbNo Then Exit Sub
    pC.Visible = False
    tC.SetFocus
End Sub
Private Sub lblPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then lblPic.BorderStyle = 1
End Sub
Private Sub lblPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblPic.BorderStyle = 0
End Sub

Private Sub tF_GotFocus(Index As Integer)
    tF(Index).ForeColor = RGB(0, 0, 0)
End Sub
Private Sub tF_LostFocus(Index As Integer)
    On Error GoTo errPath
    Select Case Index
        Case 0
            If CheckFile(tF(0).Text) <> 0 Then GoTo errPath
        Case 1
            If CheckFile(tF(1).Text) = 2 Then GoTo errPath
    End Select
    tF(Index).ForeColor = RGB(0, 0, 255)
    Exit Sub
errPath:
    tF(Index).ForeColor = RGB(255, 0, 0)
End Sub

Private Sub tF_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With tF(Index)
        .Text = Data.Files(1)
        .SelStart = Len(.Text)
    End With
    tF_LostFocus Index
End Sub

