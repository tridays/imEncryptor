Imports System.Runtime.InteropServices

Public Class cBitmap

    <StructLayoutAttribute(LayoutKind.Sequential, Pack:=1, Size:=14)> _
    Private Structure type_BMP_FileHeader
        Dim Type As Int16       '�ļ����ͱ�ʶ�� ��Ϊ��BM��
        Dim Size As Int32       '�ļ�����       ��λ:�ֽ�
        Dim Reserved As Int32   '������         �洢�����༸���ֽڵ���Ϣ
        Dim Offbits As Int32    '����ƫ����     ��λ:�ֽ�
    End Structure
    <StructLayoutAttribute(LayoutKind.Sequential, Pack:=1, Size:=40)> _
    Private Structure type_BMP_InfoHeader
        Dim Size As Int32               '��Ϣͷ����
        Dim Width As Int32              'λͼ���       ��λ:����
        Dim Height As Int32             'λͼ�߶�       ��λ:����
        Dim Planes As Int16             'Ŀ���豸����   ��Ϊ��1��
        Dim BitCount As Int16
        '����ɫλ        ֵ ����    ��ɫ��ʹ����
        '                1  ˫ɫ    2
        '                4  16ɫ    16
        '                8  256ɫ   256
        '                16 �߲�ɫ  0
        '                24 ���ɫ  0
        '                32         0
        Dim Compression As Int32        'ѹ������       0=δѹ�� 1=RLE-8 2=RLE-4
        Dim SizeImage As Int32          '���ĳ���       ��λ:�ֽ�
        Dim XPelsPerMeter As Int32      'ˮƽ�ֱ���     ��λ:��
        Dim YPelsPerMeter As Int32      '��ֱ�ֱ���     ��λ:��
        Dim ColorUsed As Int32          '��ɫ��ʹ����
        Dim ColorImportant As Int32     '��Ҫ��ɫ��
    End Structure

    '<StructLayoutAttribute(LayoutKind.Sequential, Pack:=1)> _
    Private Structure type_BMP
        '<MarshalAs(UnmanagedType.ByValArray, SizeConst:=14)> _
        Dim File As type_BMP_FileHeader
        '<MarshalAs(UnmanagedType.ByValArray, SizeConst:=40)> _
        Dim Info As type_BMP_InfoHeader
        'Dim Quad() As UInt32
        'B G R A ���� Alpha ��Ϊ 0
        '��ʹ����ɫ��
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

    Sub LoadFile(ByRef Path As String) '�� BMP �ļ�
        Initialize()

        Me.FromArrayBMP(IO.File.ReadAllBytes(Path))

        BM2 = Bitmap.FromFile(Path)
    End Sub
    Sub SaveFile(ByRef Path As String) '�� BMP �ļ�
        IO.File.WriteAllBytes(Path, Me.ToArrayBMP)
    End Sub

    Sub FromArrayBMP(ByRef B() As Byte) '���� BMP ����
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
    Function ToArrayBMP() As Byte() '��� BMP ����
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

    Sub FromArray(ByRef B() As Byte) '���� ����
        Initialize()

        Dim FLen As Long = B.Length
        Dim Count As Long = -(Int((-FLen / 3)))

        Dim PEft As Long, PLen As Long
        With BMP.Info
            Square(Count, .Width, .Height)

            PEft = PerlineEffect
            PLen = PerlineLength
            .SizeImage = PLen * .Height
            BMP.File.Size = .SizeImage + BMP.File.Offbits
        End With
        With BMP
            .File.Reserved = (3 - (FLen Mod 3)) Mod 3

            ReDim .Data(.Info.SizeImage - 1)
            Dim i As Long
            Dim Pos1 As Long, Pos2 As Long
            For i = 1 To .Info.Height
                CopyMem(.Data, Pos1, B, Pos2, PEft)
                Pos1 = Pos1 + PLen
                Pos2 = Pos2 + PEft
                Application.DoEvents()
            Next i
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
                A = A - 1
                Application.DoEvents()
            Loop
            B = s / A
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
            Return BMP.Info.Width * (BMP.Info.BitCount / 8) + 3.825 '(BMP.Info.Width * BMP.Info.BitCount + 31) / 8
        End Get
        Set(ByVal Value As Integer)
            '
        End Set
    End Property

    Function ToArray() As Byte() '��� ����
        Dim PEft As Long, PLen As Long
        PEft = PerlineEffect
        PLen = PerlineLength

        Dim B() As Byte
        Dim i As Long, Pos1 As Long, Pos2 As Long
        With BMP
            ReDim B(.Info.Height * PEft - .File.Reserved - 1)
            For i = 1 To .Info.Height - 1
                CopyMem(B, Pos2, .Data, Pos1, PEft)
                Pos1 = Pos1 + PLen
                Pos2 = Pos2 + PEft
                Application.DoEvents()
            Next i
            CopyMem(B, Pos2, .Data, Pos1, PEft - .File.Reserved)
        End With
        Return B
    End Function
    Private Sub CopyMem(ByRef Dst() As Byte, ByRef DstPos As Integer, ByRef Src() As Byte, ByRef SrcPos As Integer, ByRef Length As Integer)
        Marshal.Copy(Marshal.UnsafeAddrOfPinnedArrayElement(Src, SrcPos), Dst, DstPos, Length)
    End Sub

    Sub Bin2BMP(ByRef B() As Byte) '�� BM2 ����
        Dim s As New IO.MemoryStream
        s.Write(B, 0, B.Length)
        BM2 = Bitmap.FromStream(s)
        s.Close()
        s.Dispose()
    End Sub
    Function BMP2Bin() As Byte() '��� BM2 ����
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