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
'Private Type type_BMP_RgbQuad
'    B As Byte   '蓝色值             Blue
'    G As Byte   '绿色值             Green
'    R As Byte   '红色值             Red
'    A As Byte   '透明度 恒为“0”   Alpha
'End Type
Private Type type_BMP
    File As type_BMP_FileHeader
    Info As type_BMP_InfoHeader
   'Quad() As type_BMP_RgbQuad
    Data() As Byte
End Type

Private Declare Sub CopyMem Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim BMP As type_BMP

Private Sub Class_Initialize()
    Dim t As type_BMP
    BMP = t
    With BMP
        With .File
            .Type = &H4D42 'BM
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

Function LoadBitmap(Path As String) As Long
    Class_Initialize
    
    On Error GoTo errLoad
    Dim FreeF As Integer ', HaveQuad As Boolean
    FreeF = FreeFile
'    HaveQuad = True
    Open Path For Binary Access Read As #FreeF
    With BMP
        Get #FreeF, , .File
        Get #FreeF, , .Info
'        Select Case .Info.BitCount
'            Case 1
'                ReDim .Quad(0 To 1)
'            Case 4
'                ReDim .Quad(0 To 15)
'            Case 8
'                ReDim .Quad(0 To 255)
'            Case 16, 24, 32
'                Erase .Quad
'                HaveQuad = False
'        End Select
'        If HaveQuad Then
'            Dim i As Long
'            For i = LBound(.Quad) To UBound(.Quad)
'                Get #FreeF, , .Quad(i)
'            Next
'        End If
        ReDim .Data(0 To LOF(FreeF) - Seek(FreeF))
        Get #FreeF, , .Data
    End With
    Close #FreeF
    Exit Function
errLoad:
    LoadBitmap = 1 '读取位图文件失败
End Function
Sub SaveBitmap(Path As String)
    Dim FreeF As Integer
    FreeF = FreeFile
    Open Path For Output As #FreeF
    Close #FreeF
    Open Path For Binary Access Write As #FreeF
    '   Put #FreeF, 1, BMP '这样会有多余数据
        Put #FreeF, 1, BMP.File
        Put #FreeF, , BMP.Info
       'Put #FreeF, , BMP.Quad
        Put #FreeF, , BMP.Data
    Close #FreeF
End Sub

Sub GetDataFromArray(B() As Byte)
    Dim FLen As Long, Count As Long
    Dim PEft As Long, PLen As Long
    Dim Pos1 As Long, Pos2 As Long
        
    Class_Initialize
    
    FLen = UBound(B) + 1        '数据长度，默认数组下标为 0
    Count = -(Int((-FLen / 3))) '色块数目
    
    With BMP.Info
        Square Count, .Width, .Height
        
        PEft = PerlineEffect
        PLen = PerlineLength
        .SizeImage = PLen * .Height
        BMP.File.Size = .SizeImage + BMP.File.Offbits
        
    End With
    With BMP
        .File.Reserved = FLen
        ReDim .Data(.Info.SizeImage - 1)
        For Pos2 = 0 To FLen - PEft - 1 Step PEft
            CopyMem .Data(Pos1), B(Pos2), PEft
            Pos1 = Pos1 + PLen
           'DoEvents
        Next
        CopyMem .Data(Pos1), B(Pos2), FLen - Pos2
    End With
End Sub
Sub GetDataFromFile(Path As String)
    Dim FreeF As Integer
    Dim FLen As Long, Count As Long
    Dim PEft As Long, PLen As Long
    Dim i As Long, Pos As Long, B() As Byte
    
    Class_Initialize
    FreeF = FreeFile
    Open Path For Binary Access Read As #FreeF
    
    FLen = LOF(FreeF)
    Count = -(Int((-FLen / 3)))
    
    With BMP.Info
        Square Count, .Width, .Height
        
        PEft = PerlineEffect
        PLen = PerlineLength
        .SizeImage = PLen * .Height
        BMP.File.Size = .SizeImage + BMP.File.Offbits
        
    End With
    With BMP
        .File.Reserved = FLen
        ReDim .Data(.Info.SizeImage - 1)
        ReDim B(PEft - 1)
        For i = 0 To FLen - PEft - 1 Step PEft
            Get #FreeF, , B
            CopyMem .Data(Pos), B(0), PEft
            Pos = Pos + PLen
           'DoEvents
        Next
        ReDim B(FLen - 1 - i)
        Get #FreeF, , B
        CopyMem .Data(Pos), B(0), UBound(B) + 1
    End With
End Sub
Private Sub Square(ByRef S As Long, ByRef Width As Long, ByRef Height As Long)
    Dim t As Long, n As Long
    t = Int(Sqr(S))
    If t * t < S Then t = t + 1
    Width = t
    Height = t
End Sub

Private Property Get PerlineEffect() As Long
    PerlineEffect = BMP.Info.Width * 3
End Property
Private Property Get PerlineLength() As Long
    Dim i As Long, t As Long
    i = BMP.Info.Width * BMP.Info.BitCount \ 8
    t = i Mod 4
    PerlineLength = i + IIf(t, 4 - t, 0)
End Property

Function PutDataToFile(Path As String) As Long
    Dim FreeF As Integer
    Dim PEft As Long, PLen As Long
    Dim i As Long, Pos As Long, B() As Byte
    
    FreeF = FreeFile
    PEft = PerlineEffect
    PLen = PerlineLength
    
    On Error GoTo errFile
    Open Path For Output As #FreeF
    Close #FreeF
    
    On Error GoTo 0
    Open Path For Binary Access Write As #FreeF
    
    With BMP
        ReDim B(0 To PEft - 1)
        For i = 0 To .File.Reserved - PEft - 1 Step PEft
            CopyMem B(0), .Data(Pos), PEft
            Put #FreeF, , B
            Pos = Pos + PLen
           'DoEvents
        Next
        ReDim B(.File.Reserved - i)
        CopyMem B(0), .Data(Pos), .File.Reserved - i
        Put #FreeF, , B
    End With
    Close #FreeF
    Exit Function
errFile:
    PutDataToFile = 1 '目标路径错误或文件不可访问
End Function
Sub PutDataToArray(ByRef B() As Byte)
    Dim PEft As Long, PLen As Long
    PEft = PerlineEffect
    PLen = PerlineLength
    
    Dim Pos1 As Long, Pos2 As Long
    With BMP
        ReDim B(.File.Reserved - 1)
        For Pos2 = 0 To .File.Reserved - 1 - PEft Step PEft
            CopyMem B(Pos2), .Data(Pos1), PEft
            Pos1 = Pos1 + PLen
           'DoEvents
        Next
        CopyMem B(Pos2), .Data(Pos1), .File.Reserved - Pos2
    End With
End Sub
