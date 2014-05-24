Imports System.Runtime.InteropServices

Public Class cBitmap

    <StructLayoutAttribute(LayoutKind.Sequential, Pack:=1, Size:=14)> _
    Private Structure type_BMP_FileHeader
        Dim Type As Int16       '文件类型标识符 恒为“BM”
        Dim Size As Int32       '文件长度       单位:字节
        Dim Reserved As Int32   '保留字         存储最后多余几个字节的信息
        Dim Offbits As Int32    '正文偏移量     单位:字节
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, Pack:=1, Size:=40)> _
    Private Structure type_BMP_InfoHeader
        Dim Size As Int32               '信息头长度
        Dim Width As Int32              '位图宽度       单位:像素
        Dim Height As Int32             '位图高度       单位:像素
        Dim Planes As Int16             '目标设备级别   恒为“1”
        Dim BitCount As Int16
        '像素色位        值 名称    颜色表使用数
        '                1  双色    2
        '                4  16色    16
        '                8  256色   256
        '                16 高彩色  0
        '                24 真彩色  0
        '                32         0
        Dim Compression As Int32        '压缩类型       0=未压缩 1=RLE-8 2=RLE-4
        Dim SizeImage As Int32          '正文长度       单位:字节
        Dim XPelsPerMeter As Int32      '水平分辨率     单位:米
        Dim YPelsPerMeter As Int32      '垂直分辨率     单位:米
        Dim ColorUsed As Int32          '颜色表使用数
        Dim ColorImportant As Int32     '重要颜色数
    End Structure

    '<StructLayoutAttribute(LayoutKind.Sequential, Pack:=1)> _
    Private Structure type_BMP
        '<MarshalAs(UnmanagedType.ByValArray, SizeConst:=14)> _
        Dim File As type_BMP_FileHeader
        '<MarshalAs(UnmanagedType.ByValArray, SizeConst:=40)> _
        Dim Info As type_BMP_InfoHeader
        'Dim Quad() As UInt32
        'B G R A 其中 Alpha 恒为 0
        '不使用颜色表
        Dim Data() As Byte
    End Structure

    Dim BMP As type_BMP
    Dim BM2 As Bitmap

    Private Sub Initialize()
        Dim t As type_BMP
        With t
            With .File
                .Type = &H4D42 'BM
                .Size = 54
                .Reserved = 0
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
            'Erase .Quad
            Erase .Data
        End With
        BMP = t
    End Sub
    Sub New()
        Initialize()
    End Sub

    Sub LoadFile(ByRef Path As String) '读 BMP 文件
        Initialize()

        Me.FromArrayBMP(IO.File.ReadAllBytes(Path))

        BM2 = Bitmap.FromFile(Path)
    End Sub
    Sub SaveFile(ByRef Path As String) '存 BMP 文件
        IO.File.WriteAllBytes(Path, Me.ToArrayBMP)
    End Sub

    Sub FromArrayBMP(ByRef B() As Byte) '读入 BMP 数组
        Initialize()
        With BMP
            Dim Ptr As IntPtr
            Ptr = Marshal.UnsafeAddrOfPinnedArrayElement(B, 0)
            .File = CType(Marshal.PtrToStructure(Ptr, GetType(type_BMP_FileHeader)), type_BMP_FileHeader)
            Ptr = Marshal.UnsafeAddrOfPinnedArrayElement(B, 14)
            .Info = CType(Marshal.PtrToStructure(Ptr, GetType(type_BMP_InfoHeader)), type_BMP_InfoHeader)
            ReDim .Data(B.Length - 55)
            CopyMem(.Data, 0, B, 54, B.Length - 54)
        End With
    End Sub
    Function ToArrayBMP() As Byte() '输出 BMP 数组
        With BMP
            Dim B(53 + .Data.Length) As Byte
            CopyFromType(B, 0, .File, 14)
            CopyFromType(B, 14, .Info, 40)
            CopyMem(B, 54, .Data, 0, .Data.Length)
            Return B
        End With
    End Function
    Private Sub CopyFromType(ByRef Dst() As Byte, ByRef DstPos As Integer, ByRef Src As Object, ByRef Length As Integer)
        Dim Ptr As IntPtr = Marshal.AllocHGlobal(Length)
        Marshal.StructureToPtr(Src, Ptr, False)
        Marshal.Copy(Ptr, Dst, DstPos, Length)
        Marshal.FreeHGlobal(Ptr)
    End Sub

    Sub FromArray(ByRef B() As Byte) '读入 数据
        Initialize()

        Dim FLen As Integer = B.Length
        Dim Count As Integer = -(Int((-FLen / 3)))

        Dim PEft As Integer, PLen As Integer
        Dim t As Integer
        With BMP.Info
            t = Int(Math.Sqrt(Count))
            If t * t < FLen Then t += 1
            .Width = t
            .Height = t

            PEft = PerlineEffect
            PLen = PerlineLength
            .SizeImage = PLen * .Height
            BMP.File.Size = .SizeImage + BMP.File.Offbits
        End With
        With BMP
            .File.Reserved = FLen

            ReDim .Data(.Info.SizeImage - 1)
            Dim Pos1 As Integer, Pos2 As Integer
            For Pos2 = 0 To FLen - PEft - 1 Step PEft
                CopyMem(.Data, Pos1, B, Pos2, PEft)
                Pos1 = Pos1 + PLen
            Next
            CopyMem(.Data, Pos1, B, Pos2, FLen - Pos2)
        End With

        Me.Bin2BMP(Me.ToArrayBMP)
    End Sub
    Private Sub Square(ByRef s As Integer, ByRef B As Integer, ByRef A As Integer)
        Dim t As Double
        t = Math.Sqrt(CDbl(s))
        If InStr(t.ToString, ".") = 0 Then
            A = t
            B = t
        Else
            A = t \ 1
            Do While (s Mod A) And (A <> 1)
                A -= 1
                Application.DoEvents()
            Loop
            B = s \ A
        End If
    End Sub
    Private Property PerlineEffect() As Integer
        Get
            Return BMP.Info.Width * 3
        End Get
        Set(ByVal Value As Integer)
            '
        End Set
    End Property
    Private Property PerlineLength() As Integer
        Get
            Dim i As Integer, t As Integer
            i = BMP.Info.Width * BMP.Info.BitCount \ 8
            t = i Mod 4
            Return i + IIf(t, 4 - t, 0)
        End Get
        Set(ByVal Value As Integer)
            '
        End Set
    End Property

    Function ToArray() As Byte() '输出 数据
        Dim PEft As Integer, PLen As Integer
        PEft = PerlineEffect
        PLen = PerlineLength

        Dim B() As Byte
        Dim Pos1 As Integer, Pos2 As Integer
        With BMP
            ReDim B(.File.Reserved - 1)
            For Pos2 = 0 To .File.Reserved - 1 - PEft Step PEft
                CopyMem(B, Pos2, .Data, Pos1, PEft)
                Pos1 = Pos1 + PLen
            Next
            CopyMem(B, Pos2, .Data, Pos1, .File.Reserved - Pos2)
        End With
        Return B
    End Function
    Private Sub CopyMem(ByRef Dst() As Byte, ByRef DstPos As Integer, ByRef Src() As Byte, ByRef SrcPos As Integer, ByRef Length As Integer)
        Marshal.Copy(Marshal.UnsafeAddrOfPinnedArrayElement(Src, SrcPos), Dst, DstPos, Length)
    End Sub

    Sub Bin2BMP(ByRef B() As Byte) '读 BM2 数组
        Dim s As New IO.MemoryStream
        s.Write(B, 0, B.Length)
        BM2 = Bitmap.FromStream(s)
        s.Close()
        s.Dispose()
    End Sub
    Function BMP2Bin() As Byte() '输出 BM2 数组
        Dim s As New IO.MemoryStream
        BM2.Save(s, System.Drawing.Imaging.ImageFormat.Bmp)
        s.Flush()
        Dim B As Byte() = s.ToArray
        s.Dispose()
        Return B
    End Function

    Property Img As Image
        Get
            Return Image.FromHbitmap(BM2.GetHbitmap)
        End Get
        Set(ByVal Value As Image)

        End Set
    End Property
End Class