Attribute VB_Name = "Test_Picture"

'BMP文件头
Private Type BitmapFileHeader
    bfType As Integer '2Byte  位图类别 '标识 0,1 两个字节为 42 4D 低位在前，即 19778
    bfSize As Long '4Byte 表示BMP文件的大小
    bfReserved1 As Integer '2Byte
    bfReserved2 As Integer '2Byte
    bfOffBits As Long '4Byte 表示DIB数据区在BMP文件中的位置偏移量
End Type
Private Type BITMAPINFOHEADER
    biSize As Long '4B 指定结构体长度，为40
    biWidth As Long '4B 位图宽    '宽度 18,19,20,21 四个字节，低位在前
    biHeight As Long '4B 位图高   '高度 22,23,24,25 四个字节，低位在前
    biPlanes As Integer '2B 平面数 恒等于1
    biBitCount As Integer '2B 采用颜色位数,可以试1,2,4,8,16,24，新的可以是32
    biCompression As Long '4B 压缩方式,可以试0,1,2 其中0表示不压缩
    biSizeImage As Long '4B 实际位图数据占用的字节数
    biXPelsPerMeter As Long '4B X方向分辨率
    biYPelsPerMeter As Long '4B Y方向分辨率
    biClrUsed As Long '4B 使用的颜色数，如果为0，则表示默认值（2^颜色位数)
    biClrImportant As Long '4B 重要颜色数，如果为0，则表示所有颜色都是重要的
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
'filenumber: 文件号
'recnumber： 随机方式的文件记录号，若省略，则从当前记录开始读出数据
'Varname: 变量名，将读出的数据放入其中
    

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
'* 模 块 名：mdLSPicSize
'* 功能描述：读取图片尺寸信息(不加载图片，支持PNG)
'* 作    者：
'* 作者博客：
'* 日    期：2012-01-21 21:39
'* 版    本：V1.0.0
'***************************************************
'整行注释的为在读取图片尺寸时不需要的文件头信息

''JPEG（这个好麻烦）
'Private Type LSJPEGHeader
'    jSOI As Integer    '图像开始标识 0,1 两个字节为 FF D8 低位在前，即 -9985
'    jAPP0 As Integer    'APP0块标识 2,3 两个字节为 FF E0
'    jAPP0Length() As Byte 'jAPP0Length(1) As Byte   'APP0块标识后的长度，两个字节，高位在前
'    '  jJFIFName As Long         'JFIF标识 49(J) 48(F) 44(I) 52(F)
'    '  jJFIFVer1 As Byte         'JFIF版本
'    '  jJFIFVer2 As Byte         'JFIF版本
'    '  jJFIFVer3 As Byte         'JFIF版本
'    '  jJFIFUnit As Byte
'    '  jJFIFX As Integer
'    '  jJFIFY As Integer
'    '  jJFIFsX As Byte
'    '  jJFIFsY As Byte
'End Type
'Private Type LSJPEGChunk
'    jcType As Integer    '标识（按顺序）：APPn(0,1~15)为 FF E1~FF EF; DQT为 FF DB(-9217)
'    'SOFn(0~3)为 FF C0(-16129),FF C1(-15873),FF C2(-15617),FF C3(-15361)
'    'DHT为 FF C4(-15105); 图像数据开始为 FF DA
'    jcLength() As Byte 'jcLength(1) As Byte    '标识后的长度，两个字节，高位在前
'    '若标识为SOFn，则读取以下信息；否则按照长度跳过，读下一块
'    jBlock As Byte    '数据采样块大小 08 or 0C or 10
'    jHeight(1) As Byte    '高度两个字节，高位在前
'    jWidth(1) As Byte    '宽度两个字节，高位在前
'    '  jColorType As Byte        '颜色类型 03，后跟9字节，然后是DHT
'End Type
''PNG文件头
'Private Type LSPNGHeader
'    pType As Long    '标识 0,1,2,3 四个字节为 89 50(P) 4E(N) 47(G) 低位在前，即 1196314761
'    pType2 As Long    '标识 4,5,6,7 四个字节为 0D 0A 1A 0A
'    pIHDRLength As Long    'IHDR块标识后的长度，疑似固定 00 0D，高位在前，即 13
'    pIHDRName As Long    'IHDR块标识 49(I) 48(H) 44(D) 52(R)
'    Pwidth(3) As Byte    '宽度 16,17,18,19 四个字节，高位在前
'    Pheight(3) As Byte    '高度 20,21,22,23 四个字节，高位在前
'    '  pBitDepth As Byte
'    '  pColorType As Byte
'    '  pCompress As Byte
'    '  pFilter As Byte
'    '  pInterlace As Byte
'End Type
''GIF文件头（这个好简单）
'Private Type LSGIFHeader
'    gType1 As Long    '标识 0,1,2,3 四个字节为 47(G) 49(I) 46(F) 38(8) 低位在前，即 944130375
'    gType2 As Integer    '版本 4,5 两个字节为 7a单幅静止图像9a若干幅图像形成连续动画
'    gWidth As Integer    '宽度 6,7 两个字节，低位在前
'    gHeight As Integer    '高度 8,9 两个字节，低位在前
'End Type

'Public Function PictureSize(ByVal picPath As String, ByRef Width As Long, ByRef Height As Long) As String
'    Dim iFile As Integer
'    Dim jpg As LSJPEGHeader
'    Width = 0: Height = 0             '预输出：0 * 0
'    If picPath = "" Then PictureSize = "null": Exit Function          '文件路径为空
'    If Dir(picPath) = "" Then PictureSize = "not exist": Exit Function    '文件不存在
'    PictureSize = "error"             '预定义：出错
'    iFile = FreeFile()
'    Open picPath For Binary Access Read As #iFile
'    Get #iFile, , jpg
'    If jpg.jSOI = -9985 Then
'        Dim jpg2 As LSJPEGChunk, pass As Long
'        pass = 5 + jpg.jAPP0Length(0) * 256 + jpg.jAPP0Length(1)      '高位在前的计算方法
'        PictureSize = "JPEG error"    'JPEG分析出错
'        Do
'            Get #iFile, pass, jpg2
'            If jpg2.jcType = -16129 Or jpg2.jcType = -15873 Or jpg2.jcType = -15617 Or jpg2.jcType = -15361 Then
'                Width = jpg2.jWidth(0) * 256 + jpg2.jWidth(1)
'                Height = jpg2.jHeight(0) * 256 + jpg2.jHeight(1)
'                PictureSize = Width & "*" & Height
'                'PictureSize = "JPEG"  'JPEG分析成功
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
'        ' PictureSize = "BMP"           'BMP分析成功
'    Else
'        Dim png As LSPNGHeader
'        Get #iFile, 1, png
'        If png.pType = 1196314761 Then
'            Width = png.Pwidth(0) * 16777216 + png.Pwidth(1) * 65536 + png.Pwidth(2) * 256 + png.Pwidth(3)
'            Height = png.Pheight(0) * 16777216 + png.Pheight(1) * 65536 + png.Pheight(2) * 256 + png.Pheight(3)
'            PictureSize = Width & "*" & Height
'            'PictureSize = "PNG"       'PNG分析成功
'        ElseIf png.pType = 944130375 Then
'            Dim gif As LSGIFHeader
'            Get #iFile, 1, gif
'            Width = gif.gWidth
'            Height = gif.gHeight
'            PictureSize = Width & "*" & Height
'            'PictureSize = "GIF"       'GIF分析成功
'        Else
'            PictureSize = "unknow"    '文件类型未知
'        End If
'    End If
'    Close #iFile
'End Function
''*************************以下是测试代码
'Sub test()
'    Dim w As Long, h As Long
'    Dim f As String    '图片文件完成路径
'    Dim t As String
'    Dim Pwidth As Long, Pheight As Long
'    Dim Psize As String
'    f = "C:\1\1.jpg"  '图片文件完成路径
'    Psize = PictureSize(f, w, h)    '运行宏，w，h就是对应图片的width height  ,返回 width*height
'    If Len(Psize) > 0 Then
'        Pwidth = Val(Split(Psize, "*")(0))  '返回 图片 宽
'        Pheight = Val(Split(Psize, "*")(1))    '返回 图片 高
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
'Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long '剪贴板
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
'    '复制单元格区域图像
'    ''''''Range.CopyPicture xlScreen, xlBitmap
'    My_Screen_2
'
'    '打开剪贴板
'    OpenClipboard 0&
'    '获取剪贴板中bitmap数据的句柄
'    hBitmap = GetClipboardData(CF_BITMAP)
'    '关闭剪贴板
'    CloseClipboard
'    '初始化 GDI+
'    tSI.GdiplusVersion = 1
'    lRes = GdiplusStartup(lGDIP, tSI, 0)
'
'    If lRes = 0 Then
'        '从句柄创建 GDI+ 图像
'         lRes = GdipCreateBitmapFromHBITMAP(hBitmap, 0, lBitmap)
'        If lRes = 0 Then
'            Dim tJpgEncoder As GUID
'            Dim tParams As EncoderParameters
'
'            '初始化解码器的GUID标识
'            CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
'            '设置解码器参数
'            tParams.Count = 1
'            With tParams.Parameter ' Quality
'                '得到Quality参数的GUID标识
'                CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
'                .NumberOfValues = 1
'                .type = 4
'                .Value = VarPtr(quality)
'            End With
'
'            '保存图像
'            lRes = GdipSaveImageToFile(lBitmap, StrPtr(fileName), tJpgEncoder, tParams)
'
'            '销毁GDI+图像
'            GdipDisposeImage lBitmap
'        End If
'
'        '销毁 GDI+
'        GdiplusShutdown lGDIP
'    End If
'
'        Screen2JPG = Not lRes
'End Function
'
'Sub test()
''最后，只要用载入图片即可。
'
'Dim fileName
'fileName = "C:\1\1.jpg"
'Screen2JPG fileName
''Image.Picture = LoadPicture(fileName)
'
'End Sub


