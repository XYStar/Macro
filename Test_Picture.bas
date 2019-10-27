Attribute VB_Name = "Test_Picture"

'BMP�ļ�ͷ
Private Type BitmapFileHeader
    bfType As Integer '2Byte  λͼ��� '��ʶ 0,1 �����ֽ�Ϊ 42 4D ��λ��ǰ���� 19778
    bfSize As Long '4Byte ��ʾBMP�ļ��Ĵ�С
    bfReserved1 As Integer '2Byte
    bfReserved2 As Integer '2Byte
    bfOffBits As Long '4Byte ��ʾDIB��������BMP�ļ��е�λ��ƫ����
End Type
Private Type BITMAPINFOHEADER
    biSize As Long '4B ָ���ṹ�峤�ȣ�Ϊ40
    biWidth As Long '4B λͼ��    '��� 18,19,20,21 �ĸ��ֽڣ���λ��ǰ
    biHeight As Long '4B λͼ��   '�߶� 22,23,24,25 �ĸ��ֽڣ���λ��ǰ
    biPlanes As Integer '2B ƽ���� �����1
    biBitCount As Integer '2B ������ɫλ��,������1,2,4,8,16,24���µĿ�����32
    biCompression As Long '4B ѹ����ʽ,������0,1,2 ����0��ʾ��ѹ��
    biSizeImage As Long '4B ʵ��λͼ����ռ�õ��ֽ���
    biXPelsPerMeter As Long '4B X����ֱ���
    biYPelsPerMeter As Long '4B Y����ֱ���
    biClrUsed As Long '4B ʹ�õ���ɫ�������Ϊ0�����ʾĬ��ֵ��2^��ɫλ��)
    biClrImportant As Long '4B ��Ҫ��ɫ�������Ϊ0�����ʾ������ɫ������Ҫ��
End Type
 
Private Function Test_picture()
Dim bmp As BitmapFileHeader
Dim bmp1 As BITMAPINFOHEADER
Dim Num As Long
Dim temp As String
Dim Dstr As String
Dim bData As Byte
Dim Arr() As Byte
Dim Max

'Get [#]filenumber, [recnumber], varname
'filenumber: �ļ���
'recnumber�� �����ʽ���ļ���¼�ţ���ʡ�ԣ���ӵ�ǰ��¼��ʼ��������
'Varname: �������������������ݷ�������
    

    Path = "C:\1\1.bmp"
    iFile = FreeFile()
    Open Path For Binary Access Read As #iFile
    Get #iFile, , bmp
    Get #iFile, , bmp1
    '---
        Max = LOF(iFile)
        ReDim Arr(1 To Max)
        Get #iFile, , Arr
        Debug.Print Arr(1)
    '---
    
'    If bmp.bfType = &H4D42 Then
'        Debug.Print "True"
'    End If
    
    Debug.Print "bfType:" & bmp.bfType '0x4D42
    Debug.Print "bfSize:" & bmp.bfSize
    Debug.Print "bfReserved1:" & bmp.bfReserved1
    Debug.Print "bfReserved2:" & bmp.bfReserved2
    Debug.Print "bfOffBits:" & bmp.bfOffBits
    
    Debug.Print "------------"
    With bmp1
        Debug.Print "biSize:" & .biSize
        Debug.Print "biWidth:" & .biWidth
        Debug.Print "biHeight:" & .biHeight
        Debug.Print "biPlanes:" & .biPlanes
        Debug.Print "biBitCount:" & .biBitCount
        Debug.Print "biCompression:" & .biCompression
        Debug.Print "biSizeImage:" & .biSizeImage
        Debug.Print "biXPelsPerMeter:" & .biXPelsPerMeter
        Debug.Print "biYPelsPerMeter:" & .biYPelsPerMeter
        Debug.Print "biClrUsed:" & .biClrUsed
        Debug.Print "biClrImportant:" & .biClrImportant
    End With
    
    Close #iFile

End Function


Sub test_1()
    Debug.Print myBin(10)
End Sub

Private Function myBin(n, Optional L = 8)
    Do
        DoEvents
        myBin = n Mod 2 & myBin
        n = n \ 2
    Loop While n
    
    If Len(myBin) < L Then myBin = Right(String(L, "0") & myBin, L)
    
End Function

'***************************************************
'* ģ �� ����mdLSPicSize
'* ������������ȡͼƬ�ߴ���Ϣ(������ͼƬ��֧��PNG)
'* ��    �ߣ�
'* ���߲��ͣ�
'* ��    �ڣ�2012-01-21 21:39
'* ��    ����V1.0.0
'***************************************************
'����ע�͵�Ϊ�ڶ�ȡͼƬ�ߴ�ʱ����Ҫ���ļ�ͷ��Ϣ

''JPEG��������鷳��
'Private Type LSJPEGHeader
'    jSOI As Integer    'ͼ��ʼ��ʶ 0,1 �����ֽ�Ϊ FF D8 ��λ��ǰ���� -9985
'    jAPP0 As Integer    'APP0���ʶ 2,3 �����ֽ�Ϊ FF E0
'    jAPP0Length() As Byte 'jAPP0Length(1) As Byte   'APP0���ʶ��ĳ��ȣ������ֽڣ���λ��ǰ
'    '  jJFIFName As Long         'JFIF��ʶ 49(J) 48(F) 44(I) 52(F)
'    '  jJFIFVer1 As Byte         'JFIF�汾
'    '  jJFIFVer2 As Byte         'JFIF�汾
'    '  jJFIFVer3 As Byte         'JFIF�汾
'    '  jJFIFUnit As Byte
'    '  jJFIFX As Integer
'    '  jJFIFY As Integer
'    '  jJFIFsX As Byte
'    '  jJFIFsY As Byte
'End Type
'Private Type LSJPEGChunk
'    jcType As Integer    '��ʶ����˳�򣩣�APPn(0,1~15)Ϊ FF E1~FF EF; DQTΪ FF DB(-9217)
'    'SOFn(0~3)Ϊ FF C0(-16129),FF C1(-15873),FF C2(-15617),FF C3(-15361)
'    'DHTΪ FF C4(-15105); ͼ�����ݿ�ʼΪ FF DA
'    jcLength() As Byte 'jcLength(1) As Byte    '��ʶ��ĳ��ȣ������ֽڣ���λ��ǰ
'    '����ʶΪSOFn�����ȡ������Ϣ�������ճ�������������һ��
'    jBlock As Byte    '���ݲ������С 08 or 0C or 10
'    jHeight(1) As Byte    '�߶������ֽڣ���λ��ǰ
'    jWidth(1) As Byte    '��������ֽڣ���λ��ǰ
'    '  jColorType As Byte        '��ɫ���� 03�����9�ֽڣ�Ȼ����DHT
'End Type
''PNG�ļ�ͷ
'Private Type LSPNGHeader
'    pType As Long    '��ʶ 0,1,2,3 �ĸ��ֽ�Ϊ 89 50(P) 4E(N) 47(G) ��λ��ǰ���� 1196314761
'    pType2 As Long    '��ʶ 4,5,6,7 �ĸ��ֽ�Ϊ 0D 0A 1A 0A
'    pIHDRLength As Long    'IHDR���ʶ��ĳ��ȣ����ƹ̶� 00 0D����λ��ǰ���� 13
'    pIHDRName As Long    'IHDR���ʶ 49(I) 48(H) 44(D) 52(R)
'    Pwidth(3) As Byte    '��� 16,17,18,19 �ĸ��ֽڣ���λ��ǰ
'    Pheight(3) As Byte    '�߶� 20,21,22,23 �ĸ��ֽڣ���λ��ǰ
'    '  pBitDepth As Byte
'    '  pColorType As Byte
'    '  pCompress As Byte
'    '  pFilter As Byte
'    '  pInterlace As Byte
'End Type
''GIF�ļ�ͷ������ü򵥣�
'Private Type LSGIFHeader
'    gType1 As Long    '��ʶ 0,1,2,3 �ĸ��ֽ�Ϊ 47(G) 49(I) 46(F) 38(8) ��λ��ǰ���� 944130375
'    gType2 As Integer    '�汾 4,5 �����ֽ�Ϊ 7a������ֹͼ��9a���ɷ�ͼ���γ���������
'    gWidth As Integer    '��� 6,7 �����ֽڣ���λ��ǰ
'    gHeight As Integer    '�߶� 8,9 �����ֽڣ���λ��ǰ
'End Type

'Public Function PictureSize(ByVal picPath As String, ByRef Width As Long, ByRef Height As Long) As String
'    Dim iFile As Integer
'    Dim jpg As LSJPEGHeader
'    Width = 0: Height = 0             'Ԥ�����0 * 0
'    If picPath = "" Then PictureSize = "null": Exit Function          '�ļ�·��Ϊ��
'    If Dir(picPath) = "" Then PictureSize = "not exist": Exit Function    '�ļ�������
'    PictureSize = "error"             'Ԥ���壺����
'    iFile = FreeFile()
'    Open picPath For Binary Access Read As #iFile
'    Get #iFile, , jpg
'    If jpg.jSOI = -9985 Then
'        Dim jpg2 As LSJPEGChunk, pass As Long
'        pass = 5 + jpg.jAPP0Length(0) * 256 + jpg.jAPP0Length(1)      '��λ��ǰ�ļ��㷽��
'        PictureSize = "JPEG error"    'JPEG��������
'        Do
'            Get #iFile, pass, jpg2
'            If jpg2.jcType = -16129 Or jpg2.jcType = -15873 Or jpg2.jcType = -15617 Or jpg2.jcType = -15361 Then
'                Width = jpg2.jWidth(0) * 256 + jpg2.jWidth(1)
'                Height = jpg2.jHeight(0) * 256 + jpg2.jHeight(1)
'                PictureSize = Width & "*" & Height
'                'PictureSize = "JPEG"  'JPEG�����ɹ�
'                Stop
'                Exit Do
'            End If
'            pass = pass + jpg2.jcLength(0) * 256 + jpg2.jcLength(1) + 2
'        Loop While jpg2.jcType <> -15105    'And pass < LOF(iFile)
'    ElseIf jpg.jSOI = 19778 Then
'        Dim bmp As BitmapInfoHeader
'        Get #iFile, 15, bmp
'        Width = bmp.biWidth
'        Height = bmp.biHeight
'        PictureSize = Width & "*" & Height
'        ' PictureSize = "BMP"           'BMP�����ɹ�
'    Else
'        Dim png As LSPNGHeader
'        Get #iFile, 1, png
'        If png.pType = 1196314761 Then
'            Width = png.Pwidth(0) * 16777216 + png.Pwidth(1) * 65536 + png.Pwidth(2) * 256 + png.Pwidth(3)
'            Height = png.Pheight(0) * 16777216 + png.Pheight(1) * 65536 + png.Pheight(2) * 256 + png.Pheight(3)
'            PictureSize = Width & "*" & Height
'            'PictureSize = "PNG"       'PNG�����ɹ�
'        ElseIf png.pType = 944130375 Then
'            Dim gif As LSGIFHeader
'            Get #iFile, 1, gif
'            Width = gif.gWidth
'            Height = gif.gHeight
'            PictureSize = Width & "*" & Height
'            'PictureSize = "GIF"       'GIF�����ɹ�
'        Else
'            PictureSize = "unknow"    '�ļ�����δ֪
'        End If
'    End If
'    Close #iFile
'End Function
''*************************�����ǲ��Դ���
'Sub test()
'    Dim w As Long, h As Long
'    Dim f As String    'ͼƬ�ļ����·��
'    Dim t As String
'    Dim Pwidth As Long, Pheight As Long
'    Dim Psize As String
'    f = "C:\1\1.jpg"  'ͼƬ�ļ����·��
'    Psize = PictureSize(f, w, h)    '���к꣬w��h���Ƕ�ӦͼƬ��width height  ,���� width*height
'    If Len(Psize) > 0 Then
'        Pwidth = Val(Split(Psize, "*")(0))  '���� ͼƬ ��
'        Pheight = Val(Split(Psize, "*")(1))    '���� ͼƬ ��
'    End If
'End Sub

'Option Explicit
'
'Private Type GUID
'    Data1 As Long
'    Data2 As Integer
'    Data3 As Integer
'    Data4(0 To 7) As Byte
'End Type
'Private Type GdiplusStartupInput
'    GdiplusVersion As Long
'    DebugEventCallback As Long
'    SuppressBackgroundThread As Long
'    SuppressExternalCodecs As Long
'End Type
'Private Type EncoderParameter
'    GUID As GUID
'    NumberOfValues As Long
'    type As Long
'    Value As Long
'End Type
'Private Type EncoderParameters
'    Count As Long
'    Parameter As EncoderParameter
'End Type
'Private Declare Sub keybd_event Lib "user32" _
'(ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
'Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
'Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hpal As Long, Bitmap As Long) As Long
'Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
'Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal fileName As Long, clsidEncoder As GUID, encoderParams As Any) As Long
'Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
''Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, Bitmap As Long) As Long
'Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Integer) As Long
'Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Integer) As Long
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long '������
'Private Declare Function CloseClipboard Lib "user32" () As Long
'Const CF_BITMAP = 2
'Private Sub My_Screen_1()
'    Call keybd_event(vbKeySnapshot, 0, 0, 0)
'    DoEvents
'End Sub
'
'Private Sub My_Screen_2()
'    Call keybd_event(vbKeySnapshot, 1, 1, 1)
'    DoEvents
'End Sub
'Private Function Screen2JPG(ByVal fileName As String, Optional ByVal quality As Byte = 80) As Boolean
'
'    Dim tSI As GdiplusStartupInput
'    Dim lRes As Long
'    Dim lGDIP As Long
'    Dim lBitmap As Long
'    Dim hBitmap As Long
'    '���Ƶ�Ԫ������ͼ��
'    ''''''Range.CopyPicture xlScreen, xlBitmap
'    My_Screen_2
'
'    '�򿪼�����
'    OpenClipboard 0&
'    '��ȡ��������bitmap���ݵľ��
'    hBitmap = GetClipboardData(CF_BITMAP)
'    '�رռ�����
'    CloseClipboard
'    '��ʼ�� GDI+
'    tSI.GdiplusVersion = 1
'    lRes = GdiplusStartup(lGDIP, tSI, 0)
'
'    If lRes = 0 Then
'        '�Ӿ������ GDI+ ͼ��
'         lRes = GdipCreateBitmapFromHBITMAP(hBitmap, 0, lBitmap)
'        If lRes = 0 Then
'            Dim tJpgEncoder As GUID
'            Dim tParams As EncoderParameters
'
'            '��ʼ����������GUID��ʶ
'            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'            '���ý���������
'            tParams.Count = 1
'            With tParams.Parameter ' Quality
'                '�õ�Quality������GUID��ʶ
'                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
'                .NumberOfValues = 1
'                .type = 4
'                .Value = VarPtr(quality)
'            End With
'
'            '����ͼ��
'            lRes = GdipSaveImageToFile(lBitmap, StrPtr(fileName), tJpgEncoder, tParams)
'
'            '����GDI+ͼ��
'            GdipDisposeImage lBitmap
'        End If
'
'        '���� GDI+
'        GdiplusShutdown lGDIP
'    End If
'
'        Screen2JPG = Not lRes
'End Function
'
'Sub test()
''���ֻҪ������ͼƬ���ɡ�
'
'Dim fileName
'fileName = "C:\1\1.jpg"
'Screen2JPG fileName
''Image.Picture = LoadPicture(fileName)
'
'End Sub


