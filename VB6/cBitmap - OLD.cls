VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type type_BMP_FileHeader
    Type        As Integer  '文件类型标识符 恒为“BM”
    Size        As Long     '文件长度       单位:字节
    Reserved    As Long     '保留字
    Offbits     As Long     '正文偏移量     单位:字节
End Type
Private Type type_BMP_InfoHeader
    Size            As Long     '信息头长度
    Width           As Long     '位图宽度       单位:像素
    Height          As Long     '位图高度       单位:像素
    Planes          As Integer  '目标设备级别   恒为“1”
    BitCount        As Integer  '像素色位        值 名称    颜色表使用数
                                                '1  双色    2
                                                '4  16色    16
                                                '8  256色   256
                                                '16 高彩色  0
                                                '24 真彩色  0
                                                '32         0
    Compression     As Long     '压缩类型       0=RGB[无压缩],1=RLE8,2=RLE4,3=未压缩的16|32位图
    SizeImage       As Long     '正文长度       单位:字节
    XPelsPerMeter   As Long     '水平分辨率     单位:米
    YPelsPerMeter   As Long     '垂直分辨率     单位:米
    ColorUsed       As Long     '颜色表使用数
    ColorImportant  As Long     '重要颜色数
End Type
Private Type type_BMP_RgbQuad
    B As Byte   '蓝色值             Blue
    G As Byte   '绿色值             Green
    R As Byte   '红色值             Red
    A As Byte   '透明度 恒为“0”   Alpha
End Type
Private Type type_BMP
    File As type_BMP_FileHeader
    Info As type_BMP_InfoHeader
    Quad() As type_BMP_RgbQuad
    Data() As Byte
End Type

Private Declare Sub CopyMem Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim BMP As type_BMP

Private Sub Class_Initialize()
    Dim t As type_BMP
    BMP = t
    With BMP
        With .File
            .Type = &H4D42
            .Offbits = 54
        End With
        With .Info
            .Size = 40
            .Planes = 1
            .BitCount = 24
            .Compression = 0
            .XPelsPerMeter = 0
            .YPelsPerMeter = 0
            .ColorUsed = 0
            .ColorImportant = 0
        End With
    End With
End Sub

Function LoadFile(Path As String) As Long
    Class_Initialize
    
    On Error GoTo errLoad
    Dim FreeF As Integer
    FreeF = FreeFile
    Open Path For Binary Access Read As #FreeF
    With BMP
        Get #FreeF, , .File
        Get #FreeF, , .Info
        Select Case .Info.BitCount
            Case 1
                ReDim .Quad(0 To 1)
            Case 4
                ReDim .Quad(0 To 15)
            Case 8
                ReDim .Quad(0 To 255)
            Case 16, 24, 32
                Erase .Quad
                GoTo GetData
        End Select
        Dim i As Long
        For i = LBound(.Quad) To UBound(.Quad)
            Get #FreeF, , .Quad(i)
        Next i
GetData:
        ReDim .Data(0 To LOF(FreeF) - Seek(FreeF))
        Get #FreeF, , .Data
    End With
    Close #FreeF
    Exit Function
errLoad:
    LoadFile = 1 '文件格式错误
End Function
Sub SaveFile(Path As String)
    Dim FreeF As Integer
    FreeF = FreeFile
    Open Path For Output As #FreeF
    Close #FreeF
    Open Path For Binary Access Write As #FreeF
    '   Put #FreeF, 1, BMP '这样会有多余数据
        Put #FreeF, 1, BMP.File
        Put #FreeF, , BMP.Info
        Put #FreeF, , BMP.Quad
        Put #FreeF, , BMP.Data
    Close #FreeF
End Sub

Sub LoadDataEx(B() As Byte)
    Class_Initialize
    
    Dim FLen As Long
    FLen = UBound(B) - LBound(B) + 1
    Dim Count As Long
    Count = -(Int((-FLen / 3)))
    
    With BMP.Info
        Square Count, .Width, .Height
        
        Dim PEft As Long, PLen As Long
        PEft = PerlineEffect
        PLen = PerlineLength
        .SizeImage = PLen * .Height
        BMP.File.Size = .SizeImage + BMP.File.Offbits
        
    End With
    With BMP
        .File.Reserved = (3 - (FLen Mod 3)) Mod 3
        Erase .Quad
        ReDim .Data(0 To .Info.SizeImage - 1)
        Dim i As Long, Pos1 As Long, Pos2 As Long
        Pos1 = LBound(.Data): Pos2 = LBound(B)
        For i = 1 To .Info.Height
            CopyMem .Data(Pos1), B(Pos2), PEft
            Pos1 = Pos1 + PLen
            Pos2 = Pos2 + PEft
            DoEvents
        Next i
    End With
End Sub
Sub LoadData(Path As String)
    Class_Initialize
    
    Dim FreeF As Integer
    FreeF = FreeFile
    Open Path For Binary Access Read As #FreeF
    
    Dim FLen As Long
    FLen = LOF(FreeF)
    Dim Count As Long
    Count = -(Int((-FLen / 3)))
    
    With BMP.Info
        Square Count, .Width, .Height
        
        Dim PEft As Long, PLen As Long
        PEft = PerlineEffect
        PLen = PerlineLength
        .SizeImage = PLen * .Height
        BMP.File.Size = .SizeImage + BMP.File.Offbits
        
    End With
    With BMP
        .File.Reserved = FLen Mod 3
        Erase .Quad
        ReDim .Data(0 To .Info.SizeImage - 1)
        Dim i As Long, Pos As Long, B() As Byte
        Pos = LBound(.Data): ReDim B(0 To PEft - 1)
        For i = 1 To .Info.Height
            Get #FreeF, , B
            CopyMem .Data(Pos), B(0), PEft
            Pos = Pos + PLen
            DoEvents
        Next i
    End With
End Sub
Private Sub Square(S, ByRef B As Long, ByRef A As Long)
    Dim Temp
    Temp = Sqr(S)
    If InStr(Temp, ".") = 0 Then
        A = Temp: B = Temp
    Else
        A = Int(Temp)
        Do While (S Mod A)
            If A = 1 Then Exit Do
            A = A - 1
            DoEvents
        Loop
        B = S / A
    End If
End Sub

Private Property Get PerlineEffect() As Long
    PerlineEffect = BMP.Info.Width * 3
End Property
Private Property Get PerlineLength() As Long
    'PerlineLength = (BMP.Info.Width * BMP.Info.BitCount + 31) / 8
    PerlineLength = BMP.Info.Width * (BMP.Info.BitCount / 8) + 3.825
End Property

Sub SaveData(Path As String)
    Dim PEft As Long, PLen As Long
    PEft = PerlineEffect
    PLen = PerlineLength
    
    Dim FreeF As Integer
    FreeF = FreeFile
    Open Path For Output As #FreeF
    Close #FreeF
    Open Path For Binary Access Write As #FreeF
    
    Dim i As Long, Pos As Long, B() As Byte
    With BMP
        Pos = LBound(.Data): ReDim B(0 To PEft - 1)
        For i = 1 To .Info.Height - 1
            CopyMem B(0), .Data(Pos), PEft
            Put #FreeF, , B
            Pos = Pos + PLen
            'Debug.Assert (.Data(Pos - 1) = 0)
            DoEvents
        Next i
        ReDim B(0 To PEft - 1 - .File.Reserved)
        CopyMem B(0), .Data(Pos), PEft - .File.Reserved
        Put #FreeF, , B
    End With
    Close #FreeF
End Sub
Sub SaveDataEx(B() As Byte)
    Dim PEft As Long, PLen As Long
    PEft = PerlineEffect
    PLen = PerlineLength
    
    Dim i As Long, Pos1 As Long, Pos2 As Long
    With BMP
        Pos1 = LBound(.Data): ReDim B(0 To (.Info.Height * PEft) - .File.Reserved - 1)
        For i = 1 To .Info.Height - 1
            CopyMem B(Pos2), .Data(Pos1), PEft
            Pos1 = Pos1 + PLen
            Pos2 = Pos2 + PEft
            DoEvents
        Next i
        CopyMem B(Pos2), .Data(Pos1), PEft - .File.Reserved
    End With
End Sub
