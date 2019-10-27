Attribute VB_Name = "Test_Picture_Enforce"
 
'2016-09-02
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
 
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, _
                                                                      RefIID As GUID, _
                                                                      ByVal fPictureOwnsHandle As Long, _
                                                                      IPic As IPicture) As Long

Private Const CF_BITMAP = 2
Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hpal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

 
'====
'2016-08-26
'SelectObject:将选择的一个对象放到指定的设备环境中,新对象替换先前的相同类型的对象
'在当前设备场景选择图像对象
'hdc为设备环境句柄
'hObject为被选择对象的句柄，该对象可以是位图等，且必须由指定的函数建立
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
 
Private Type Bitmap '14 bytes
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbAlpha As Byte
End Type

Private Type BITMAPINFOHEADER
   bmSize As Long
   bmWidth As Long
   bmHeight As Long
   bmPlanes As Integer
   bmBitCount As Integer
   bmCompression As Long
   bmSizeImage As Long
   bmXPelsPerMeter As Long
   bmYPelsPerMeter As Long
   bmClrUsed As Long
   bmClrImportant As Long
End Type

Private Type BITMAPINFO
   bmHeader As BITMAPINFOHEADER
   bmColors(0 To 255) As RGBQUAD
End Type
'2016-09-07 Add
'===============
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                     ByVal lpDeviceName As String, _
                                                                     ByVal lpOutput As String, _
                                                                     lpInitData As DEVMODE) As Long
                                                                     
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, _
                                                                 ByVal nWidth As Long, _
                                                                 ByVal nHeight As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, _
                                                 ByVal x As Long, _
                                                 ByVal y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, _
                                                 ByVal xSrc As Long, _
                                                 ByVal ySrc As Long, _
                                                 ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, _
                                                 ByVal dwRop As Long) As Long
                                                 
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, _
                                                        ByVal nStretchMode As Long) As Long
                                                        
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal dWidth As Long, _
                                                    ByVal dHeight As Long, _
                                                    ByVal SrcX As Long, _
                                                    ByVal SrcY As Long, _
                                                    ByVal SrcWidth As Long, _
                                                    ByVal SrcHeight As Long, _
                                                    lpBits As Any, _
                                                    lpBI As BITMAPINFO, _
                                                    ByVal wUsage As Long, _
                                                    ByVal RasterOp As Long) As Long
  
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Private Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Private Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Private Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Private Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Private Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Private Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Private Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const COLORONCOLOR = 3

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
'===============
Private Declare Function InvalidateRectAsAny _
                Lib "user32" _
                Alias "InvalidateRect" (ByVal hwnd As Long, _
                                        lpRect As Any, _
                                        ByVal bErase As Long) As Long
                                        
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Const DIB_RGB_COLORS = 0
Private Const nOffset = 1E-07

Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hDC As Long, _
                                                     lpInfoHeader As BITMAPINFOHEADER, _
                                                     ByVal dwUsage As Long, _
                                                     lpInitBits As Any, _
                                                     lpInitInfo As BITMAPINFO, _
                                                     ByVal wUsage As Long) As Long
' constants for CreateDIBitmap
Private Const CBM_CREATEDIB = &H2      '  create DIB bitmap
Private Const CBM_INIT = &H4           '  initialize bitmap

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'2016-09-02
Private Sub test_Enforce_BMP()
    
    Enforce_BMP "C:\1\1.bmp"
    
End Sub
Private Function SetMe2(tempVar)   'R
    
    If tempVar > 255 Then tempVar = 255
    If tempVar < 0 Then tempVar = 0
    SetMe2 = tempVar
 
End Function
Private Function SetMe1(tempVar) 'G
    
    If tempVar > 255 Then tempVar = 255
    If tempVar < 0 Then tempVar = 0
    SetMe1 = tempVar
    
End Function
Private Function SetMe0(tempVar) 'B

    If tempVar > 255 Then tempVar = 255
    If tempVar < 0 Then tempVar = 0
    SetMe0 = tempVar
    
End Function
'2016-09-01
Private Function Enforce_BMP() '(Path)
Dim bmp As Bitmap
Dim res As Long
Dim Picture As New StdPicture
'Dim Path
Dim Brightness
Dim myNum
Dim sum
Dim x As Long, y As Long
Dim R, G, B, n, m
Dim i, j, k, t, L, v
Dim temp
Dim W, H, WxH
Dim Flag
Const nOffset = 1E-07
'------
Dim SetFlag, GrayFlag, WriteFlag
SetFlag = 0
GrayFlag = 0
'===============
'
    Dim hDC As Long
    Dim hwnd As Long
    Dim dm As DEVMODE
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim BMI As BITMAPINFO
'===============

    myNum = 100
    Brightness = CSng(myNum) / 100
    Dim bTable0(0 To 255), bTable1(0 To 255), bTable2(0 To 255)
    Dim tempColor
    '---R
    For i = 0 To 255
        tempColor = Int(CSng(i) * Brightness)
        SetMe2 tempColor
        bTable2(i) = tempColor
    Next i
    '---G
    For i = 0 To 255
        tempColor = Int(CSng(i) * Brightness)
        SetMe1 tempColor
        bTable1(i) = tempColor
    Next i
    '---B
    For i = 0 To 255
        tempColor = Int(CSng(i) * Brightness)
        SetMe0 tempColor
        bTable0(i) = tempColor
    Next i
    '---
'===============


        '==============
'
            'Path = "C:\1\1.bmp"
            Path = "C:\1\1-1.bmp"
            'Path = "C:\1\1-2.bmp"
            'Path = "C:\1\1-3.bmp"
            'Path = "C:\1\2-1.bmp"
            'Path = "C:\1\2-2.bmp"
            'Path = "C:\1\2-3.bmp"
            'Path = "C:\1\2-4.bmp"
            Path = "C:\1\test.bmp"
            'Path = "C:\1\test2.bmp"
            'Path = "C:\1\test3.bmp"
            'Path = "C:\1\test4.bmp"
            'Path = "C:\1\2.bmp"
            Set Picture = LoadPicture(Path)
            Call GetObject(Picture.Handle, Len(bmp), bmp)
            BMPHandle = Picture.Handle
            W = bmp.bmWidth
            H = bmp.bmHeight
            WxH = W * H
            If W Mod 4 <> 0 Then

                MsgBox "BMP的宽度必须为4的倍数：" & W
                End

            End If
        '==============
        
            '==============
                
'                W = 220
'                H = 30
'
'                srcDC = CreateDC("DISPLAY", "", "", dm) '创建设备上下文环境
'                BMPHandle = CreateCompatibleBitmap(srcDC, W, H)
'                Call GetObject(BMPHandle, Len(bmp), bmp)
'                W = bmp.bmWidth
'                H = bmp.bmHeight
'                WxH = W * H
'                If W Mod 4 <> 0 Then
'
'                    MsgBox "BMP的宽度必须为4的倍数：" & W
'                    End
'
'                End If
'
            '==============
   
   
        '=============
            Dim ImageData() As Byte
            '------
            'bmp.bmWidth须为4的倍数
            ReDim ImageData(0 To 2, 0 To bmp.bmWidth - 1, 0 To bmp.bmHeight - 1)
            ''ImageData(0,0,0)
            ''The first dimension: red(2),green(1),blue(0)
            ''The second dimension will be used to address the x coordinates of the image's pixels.
            ''The second dimension will be used to address the y coordinates of the image's pixels.
            'Arrays supplied to GetBitmapBits (and later in this tutorial,
            'GetDIBits) must have a width that is a multiple of 4, e.g. 4, 8, 16, 256, 360, etc.
            'If your image has a width that is a multiple of four,
            'no worries – but if it is not a multiple of four, you will need to adjust your code accordingly.
            GetBitmapBits Picture.Handle, bmp.bmWidthBytes * bmp.bmHeight, ImageData(0, 0, 0)
            '------
        '=============

'=====
    '灰度值处理
    Dim Gray(), tGray, tempGray
    ReDim Gray(bmp.bmWidth - 1, bmp.bmHeight - 1)
    For y = 0 To bmp.bmHeight - 1
        For x = 0 To bmp.bmWidth - 1
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            Gray(x, y) = Int(R * 0.3 + G * 0.59 + B * 0.11)
            tGray = Gray(x, y)
        Next x
    Next y

'=====

'WriteFlag = 1
'SetFlag = 6
GrayFlag = 1
'Sobel边缘检测
If SetFlag = 6 Then
'============
'X
'-1 -2 -1
' 0  0  0
' 1  2  2
'Y
'-1  0  1
'-2  0  0
'-1  0  2
'(-1,-1) (0,-1) (1,-1)
'(-1, 0) (0, 0) (1, 0)
'(-1, 1) (0, 1) (1, 1)


    Dim nScale
    Dim TX, TY, tG
    nScale = 0.8
    
    tempGray = Gray
    For y = 0 To H - 1
        For x = 0 To W - 1
            If x = 0 Or x = W - 1 Or y = 0 Or y = H - 1 Then
                Gray(x, y) = 0
            Else
                '相关
                'TX = tempGray(x - 1, y + 1) + 2 * tempGray(x, y + 1) + tempGray(x + 1, y + 1) - tempGray(x - 1, y - 1) - 2 * tempGray(x, y - 1) - tempGray(x + 1, y - 1)
                'TY = tempGray(x + 1, y - 1) + 2 * tempGray(x + 1, y) + tempGray(x + 1, y + 1) - tempGray(x - 1, y - 1) - 2 * tempGray(x - 1, y) - tempGray(x - 1, y + 1)
                '卷积
                TX = tempGray(x - 1, y - 1) + 2 * tempGray(x, y - 1) + tempGray(x + 1, y - 1) - tempGray(x - 1, y + 1) - 2 * tempGray(x, y + 1) - tempGray(x + 1, y + 1)
                TY = tempGray(x - 1, y - 1) + 2 * tempGray(x - 1, y) + tempGray(x - 1, y + 1) - tempGray(x + 1, y - 1) - 2 * tempGray(x + 1, y) - tempGray(x + 1, y + 1)
                                
                tG = TX ^ 2 + TY ^ 2
                tG = CInt(Sqr(tG) + nOffset)
                
                If tG > 100 Then
                    Gray(x, y) = 255
                Else
                    Gray(x, y) = 0
                End If
                
                
            End If
        Next x
    Next y
    
     
'============
End If


'SetFlag = 5
'GrayFlag = 1
'字符分隔
If SetFlag = 5 Then
'============
    
    '扫描每行点数
    Dim rowNum(), nLeft, nRight, cMin, cMax
    ReDim rowNum(H - 1)
    '----
    Flag = 0
    For y = 0 To H - 1
        For x = 0 To W - 1
            If Gray(x, y) <> 255 Then
                rowNum(y) = rowNum(y) + 1
                cMax = x
            End If
            
            If rowNum(y) = 1 Then
                If Flag = 0 Then
                    nLeft = x
                    Flag = 1
                End If
                cMin = x '1次
            End If
        Next x
        If nLeft > cMin Then nLeft = cMin '左边界
        If nRight < cMax Then nRight = cMax '右边界
    Next y
    '----
    
    Dim myrow
    Dim rowUp(), rowDown()
    Flag = 0
    myrow = 1
    For y = 0 To H - 1
           
        If Flag = 0 Then
            If rowNum(y) > 10 Then
                ReDim Preserve rowUp(myrow)
                rowUp(myrow) = y '上边界
                Flag = 1
            End If
        End If
        
        If Flag = 1 Then
            If rowNum(y) = 0 Then
                ReDim Preserve rowDown(myrow)
                rowDown(myrow) = y '下边界
                Flag = 0
                myrow = myrow + 1
            End If
        End If
    Next y
    
    
    '先边缘检测，然后再使用模块切割
    
    '14*6 14*10 14*12
'    For i = 1 To UBound(rowUp)
'        Debug.Print rowUp(i), rowDown(i)
'    Next i
'
    For x = 0 To W - 1
        For y = 0 To H - 1
            If Gray(x, y) >= 210 Then
                Gray(x, y) = 255
            End If
        Next y
    Next x
    
'============
End If


'SetFlag = 4
'GrayFlag = 1
'自己设计的算法 0905
If SetFlag = 4 Then
'============
    Dim NowNum, NextNum, MidNum, tempNum
    tempGray = Gray
    
    sum = 0
    t = 0
    For x = 0 To W - 1
        For y = 0 To H - 1
            NowNum = Gray(x, y)
            If x + 1 > W - 1 Then
                NextNum = Gray(x, y)
            Else
                NextNum = Gray(x + 1, y)
            End If
            
            If NowNum <> 255 And NextNum <> 255 Then
                tempNum = Abs(NowNum - NextNum)
                sum = tempNum + sum
                t = t + 1
            End If
        Next y
    Next x
    
    MidNum = CInt(sum / t + nOffset)
    
    For x = 0 To W - 1
        For y = 0 To H - 1
            NowNum = Gray(x, y)
            If x + 1 > W - 1 Then
                NextNum = Gray(x, y)
            Else
                NextNum = Gray(x + 1, y)
            End If
            If NowNum <> 255 And NextNum <> 255 Then
                If Abs(NowNum - NextNum) > MidNum Then
                    If NowNum > NextNum Then
                        For i = x + 1 To W - 1
                            If Gray(i, y) = NextNum Then
                                tempGray(i, y) = 0
                            Else
                                Exit For
                            End If
                        Next i
                    Else
                        For i = x To 0 Step -1
                            If Gray(i, y) = NowNum Then
                                tempGray(i, y) = 0
                            Else
                                Exit For
                            End If
                        Next i
                    End If
                End If
            End If
        Next y
    Next x
    
    Gray = tempGray
    For x = 0 To W - 1
        For y = 0 To H - 1
            
            If Gray(x, y) < 160 Then
                Gray(x, y) = 0
            Else
                'Gray(x, y) = 255
            End If
        Next y
    Next x
    
'============
End If


'SetFlag = 3
'GrayFlag = 1
'Laplace图像锐化
If SetFlag = 3 Then
'============
'拉普拉斯算子
'---
'0  1  0
'1 -4  1
'0  1  0
'---
'1  1  1
'1 -8  1
'1  1  1
'---
    
    '---
    '锐化模板
    Dim Template
    Dim ptr
    Dim dst
    
    ReDim dst(W - 1, H - 1)
    
    '如果使用的模板中心是负数，那么必须将原图像减去经拉普拉斯变换后的图像。
    '如果使用的模板中心是正数，那么就将原图像加上经过拉普拉斯变换后的图像。
    'Template = Array(0, 1, 0, 1, -4, 1, 0, 1, 0)
    'Template = Array(1, 1, 1, 1, -8, 1, 1, 1, 1)
    Template = Array(0, -1, 0, -1, 4, -1, 0, -1, 0)
    'Template = Array(-1, -1, -1, -1, 8, -1, -1, -1, -1)
    '---
    
    '---
    '锐化滤波
        For x = 0 To W - 1
            For y = 0 To H - 1
                
                '边界处理
                If x = 0 Or x = W - 1 Or y = 0 Or y = H - 1 Then
                    dst(x, y) = Gray(x, y)
                Else
                    
                    If Gray(x, y) = 255 Then
                        x = x
                    End If
                    
                    dst(x, y) = 0
                    sum = 0
                    For m = -1 To 1
                        For n = -1 To 1
                            ptr = Gray(x + m, y + n)
                            sum = sum + ptr * Template(3 * (m + 1) + n + 1)
                        Next
                    Next
                    '-------
                        If sum < 0 Then sum = 0
                        temp = sum + Gray(x, y)
                        If temp > 255 Then temp = 255
                        dst(x, y) = temp
                    '-------
                End If
                
            Next
        Next
    '---
 
    
        For x = 0 To W - 1
            For y = 0 To H - 1
                Gray(x, y) = dst(x, y)
            Next
        Next
'============
End If


'GrayFlag = 1
'SetFlag = 2
If SetFlag = 2 Then
'图像增强,直方图均衡化
'============
Dim Num(0 To 255), P(0 To 255), P1(0 To 255)

    '存放图像各个灰度级出现的次数
    For x = 0 To W - 1
        For y = 0 To H - 1
            temp = Gray(x, y)
            Num(temp) = Num(temp) + 1
        Next
    Next
    
    '存放图像各个灰度级的出现概率
    For i = 0 To 255
        P(i) = Num(i) / WxH
    Next
    
    '存放各个灰度级之前的概率和
    For i = 0 To 255
        For k = 0 To i
            P1(i) = P1(i) + P(k)
        Next
    Next
    
    '直方图变换
    
    For x = 0 To W - 1
        For y = 0 To H - 1
            v = Gray(x, y)
            If v = 0 Then
                v = v
            End If
            temp = CInt(P1(v) * 255 + nOffset)
            Gray(x, y) = temp
        Next
    Next
'============
End If


If SetFlag = 1 Then
'###################
        '=======
        Dim mysum
        Dim Q, Q2, f, C, D, Amount
        Dim x0, y0
        Dim runFlag
        n = 2
        Flag = 0
        sum = 0
        runFlag = 0



        For x = 0 To W - 1
            For y = 0 To H - 1
'                x0 = CInt(x + (2 * n + 1) / 2 + nOffset)
'                y0 = CInt(y + (2 * n + 1) / 2 + nOffset)
                x0 = x
                y0 = y
 
                    
                    If y0 - n < 0 Or x0 - n < 0 Or x0 + n > W - 1 Or y0 + n > H - 1 Then
                        runFlag = 0
                    Else
                        runFlag = 1
                    End If
                    
If runFlag = 1 Then

                    '局部平均值m
                    '----
                    mysum = 0
                    For k = x0 - n To x0 + n
                        For L = y0 - n To y0 + n
                            mysum = mysum + Gray(k, L)
                        Next L
                    Next k
                    m = mysum / (2 * n + 1) ^ 2
                    '----
                    '局部方差Q
                    '----
                    mysum = 0
                    For k = x0 - n To x0 + n
                        For L = y0 - n To y0 + n
                            mysum = mysum + (Gray(k, L) - m) ^ 2
                        Next L
                    Next k
                    Q2 = (mysum / (2 * n + 1) ^ 2)
                    Q = Sqr(Q2)
                    '----
                    
    If Q = 0 Then
    
    Else
                    
                    '---
        C = 3
        Amount = 100
        D = Amount / Q
        GrayFlag = 0
        
        If GrayFlag = 0 Then
                        '-----
                        For k = x0 - n To x0 + n
                            For L = y0 - n To y0 + n

                                If k = 16 And L = 8 Then
                                    k = k
                                End If

                                f = CLng(m + (ImageData(2, x, y) - m) * D + nOffset)
                                If f > 255 Then
                                    ImageData(2, x, y) = 255
                                ElseIf f < 0 Then
                                    ImageData(2, x, y) = 0
                                Else
                                    ImageData(2, x, y) = f
                                End If

                                f = CLng(m + (ImageData(1, x, y) - m) * D + nOffset)
                                If f > 255 Then
                                    ImageData(1, x, y) = 255
                                ElseIf f < 0 Then
                                    ImageData(1, x, y) = 0
                                Else
                                    ImageData(1, x, y) = f
                                End If

                                f = CLng(m + (ImageData(0, x, y) - m) * D + nOffset)
                                If f > 255 Then
                                    ImageData(0, x, y) = 255
                                ElseIf f < 0 Then
                                    ImageData(0, x, y) = 0
                                Else
                                    ImageData(0, x, y) = f
                                End If
                            Next L
                        Next k

        Else
                        For k = x0 - n To x0 + n
                            For L = y0 - n To y0 + n

                                If k = 16 And L = 8 Then
                                    k = k
                                End If

                                f = CLng(m + (Gray(k, L) - m) * D + nOffset)
                                If f > 255 Then
                                    Gray(k, L) = 255
                                ElseIf f < 0 Then
                                    Gray(k, L) = 0
                                Else
                                    Gray(k, L) = CInt(f + nOffset)
                                End If
                            Next L
                        Next k
        End If
        
                        '-----
                    '---
        
                    Flag = Flag + 1
    End If 'Q=0
End If
            
            DoEvents
            
            Next y
         Next x
        '=======
'###################
End If 'SetFlag=1
      
      
      
      
'Show
'##############################################################
n = 1
If n = 1 Then
    t = 2
    For y = 0 To bmp.bmHeight - 1
        For x = 0 To bmp.bmWidth - 1
            '---
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            '---
            
            '=============
            '灰度值处理
            If GrayFlag = 1 Then
                tGray = Gray(x, y)
                R = tGray: G = tGray: B = tGray
            End If
            '=============
            
            '显示在Excel
            '=====
            If WriteFlag = 1 Then
                If x > 0 And y > 0 And x < W - 1 And y < H - 1 Then
                    Cells(y + 2, x + 1) = ""
                    Cells(y + 2, x + 1).Interior.Color = xlNone
                    If R <> 255 Then
                        Cells(y + 2, x + 1) = R
                        Cells(y + 2, x + 1).Interior.Color = RGB(R, R, R)
                    End If
                End If
            End If
            '=====
            
            '============
            '赋值
                '----
                'R
                ImageData(2, x, y) = R
                'G
                ImageData(1, x, y) = G
                'B
                ImageData(0, x, y) = B
                '---
            '============
        Next x
        t = t + 1
    Next y

End If

    
    SetBitmapBits Picture.Handle, bmp.bmWidthBytes * bmp.bmHeight, ImageData(0, 0, 0)

            
            '=============
                
'
'                MultipleW = 2
'                MultipleH = 2

                
'                trgDC = CreateCompatibleDC(0)
'                BMPHandle = Picture.Handle
'                Call SelectObject(trgDC, BMPHandle)
                
                
                'hwnd = WindowFromPoint(0, 0)
                'trgDC = GetDC(hwnd)
                'Call DeleteDC(trgDC) '删除内存环境
                'BMPHandle = Picture.Handle
                'trgDC = GetDC(0)
                'trgDC = GetWindowDC(BMPHandle)
                
 
                'trgDC = GetDC(BMPHandle)
                'Call DeleteDC(trgDC) '删除内存环境
                'Call ReleaseDC(BMPHandle, srcDC) '释放设备环境句柄
'                'srcDC = GetDC(Picture.Handle)
'                trgDC = CreateCompatibleDC(srcDC) '创建与设备相关的内存环境
'
'                'BMPHandle = Picture.Handle
'                Call SelectObject(trgDC, BMPHandle) '选择对象
'                'Call StretchBlt(trgDC, 0, 0, W * MultipleW, H * MultipleH, srcDC, 0, 0, W, H, SRCCOPY)
'                'SetBitmapBits BMPHandle, bmp.bmWidthBytes * bmp.bmHeight, ImageData(0, 0, 0)
'
'                '----
'                BMI.bmHeader.bmSize = 40
'                BMI.bmHeader.bmPlanes = 1
'                BMI.bmHeader.bmBitCount = 24
'                BMI.bmHeader.bmCompression = 0
'                BMI.bmHeader.bmWidth = W
'                BMI.bmHeader.bmHeight = H
'                Call StretchDIBits(trgDC, 0, 0, W, H, 0, 0, W, H, ImageData(0, 0, 0), BMI, 0, SRCCOPY)
                '----
            '=============
        '=============
            Call OpenClipboard(0&) '打开剪切板
            Call EmptyClipboard '清除当前剪切板中的内容
            Call SetClipboardData(CF_BITMAP, BMPHandle) '放置多个不同格式的数据项
            Call CloseClipboard '关闭剪切板
        '=============
        
 
 
            '=============
                
                If srcDC <> 0 Then Call ReleaseDC(BMPHandle, srcDC)    '释放设备环境句柄
                If trgDC <> 0 Then Call DeleteDC(trgDC) '删除内存环境
            '=============
'========
    Dim SavePath
    SavePath = "C:\1\2.bmp"
    SavePicture ApiGetClipBmp, SavePath
    'SavePicture ApiGetClipBmp, Path
'========
    
'##############################################################
  
End Function

Private Function Enforce_BMP_Test() '(Path)
Dim bmp As Bitmap
Dim res As Long
Dim Picture As New StdPicture
'Dim Path
Dim Brightness
Dim myNum
Dim sum
Dim x As Long, y As Long
Dim R, G, B, n, m
Dim i, j, k, t, L, v
Dim temp
Dim W, H, WxH
Dim Flag
Const nOffset = 1E-07
'------
Dim SetFlag, GrayFlag, WriteFlag
SetFlag = 0
GrayFlag = 0
'===============
'
    Dim hDC As Long
    Dim hwnd As Long
    Dim dm As DEVMODE
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim BMI As BITMAPINFO
'===============
 

''==============
    Path = "C:\1\6.bmp"
    Set Picture = LoadPicture(Path)
    Call GetObject(Picture.Handle, Len(bmp), bmp)
    BMPHandle = Picture.Handle
    W = bmp.bmWidth
    H = bmp.bmHeight
    WxH = W * H
''==============


'==============
    'Dim mW, mH
    'mW = 110
    'mH = 15
    srcDC = CreateDC("DISPLAY", "", "", dm) '创建设备上下文环境
    'BMPHandle = CreateCompatibleBitmap(srcDC, W, H)
    
    '----
    srcDC = 0
    trgDC = CreateCompatibleDC(0)
    Call SelectObject(trgDC, BMPHandle) '选择一对象到指定的设备上下文环境中
    '抓屏
    Call StretchBlt(trgDC, 0, 0, W * 2, H * 2, srcDC, 0, 0, W, H, SRCCOPY)
    '----

    Call GetObject(BMPHandle, Len(bmp), bmp)
    W = bmp.bmWidth
    H = bmp.bmHeight
    WxH = W * H
'==============

    BMI.bmHeader.bmSize = 40
    BMI.bmHeader.bmPlanes = 1
    BMI.bmHeader.bmBitCount = 24
    BMI.bmHeader.bmCompression = 0
    BMI.bmHeader.bmWidth = W
    BMI.bmHeader.bmHeight = H
    
'=============
    Dim ImageData() As Byte
    '------
    'bmp.bmWidth须为4的倍数
    ReDim ImageData(0 To 2, 0 To W - 1, 0 To H - 1)
    GetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)
    '------

'=====
    '灰度值处理
    Dim Gray(), tGray, tempGray
    ReDim Gray(W - 1, H - 1)
    For y = 0 To H - 1
        For x = 0 To W - 1
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            Gray(x, y) = Int(R * 0.3 + G * 0.59 + B * 0.11)
        Next x
    Next y

'=====

'Call Laplace(W, H, Gray)
GrayFlag = 1
n = 1
If n = 1 Then
    t = 2
    For y = 0 To H - 1
        For x = 0 To W - 1
            '---
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            '---
            
            '=============
            '灰度值处理
            If GrayFlag = 1 Then
                tGray = Gray(x, y)
                R = tGray: G = tGray: B = tGray
            End If
            '=============
            
            '显示在Excel
            '=====
            If WriteFlag = 1 Then
                If x > 0 And y > 0 And x < W - 1 And y < H - 1 Then
                    Cells(y + 2, x + 1) = ""
                    Cells(y + 2, x + 1).Interior.Color = xlNone
                    If R <> 255 Then
                        Cells(y + 2, x + 1) = R
                        Cells(y + 2, x + 1).Interior.Color = RGB(R, R, R)
                    End If
                End If
            End If
            '=====
            
            '============
            '赋值
                '----
                'R
                ImageData(2, x, y) = R
                'G
                ImageData(1, x, y) = G
                'B
                ImageData(0, x, y) = B
                '---
            '============
        Next x
        t = t + 1
    Next y

End If
'=============
 
    'Call SetStretchBltMode(trgDC, COLORONCOLOR)
    'Call StretchDIBits(trgDC, 0, H - 1, W * 0.5, -H * 0.5, 0, 0, W, H, ImageData(0, 0, 0), BMI, 0, SRCCOPY)
    'SetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)
    'Call StretchBlt(trgDC, 0, 0, W * 3, H * 3, srcDC, 0, 0, W, H, SRCCOPY)
'=============



        '=============
            Call OpenClipboard(0&) '打开剪切板
            Call EmptyClipboard '清除当前剪切板中的内容
            Call SetClipboardData(CF_BITMAP, BMPHandle) '放置多个不同格式的数据项
            Call CloseClipboard '关闭剪切板
        '=============
        
 

'=============
    
    If srcDC <> 0 Then Call ReleaseDC(BMPHandle, srcDC)    '释放设备环境句柄
    If trgDC <> 0 Then Call DeleteDC(trgDC) '删除内存环境
'=============
'========
    Dim SavePath
    SavePath = "C:\1\2.bmp"
    SavePicture ApiGetClipBmp, SavePath
    'SavePicture ApiGetClipBmp, Path
'========
    
'##############################################################
  
End Function
Private Function ApiGetClipBmp() As IPicture '把剪切板图片数据转换为图片
 
    Dim Pic As PicBmp
    Dim IID_IDispatch As GUID
    
    Call OpenClipboard(0&)  'OpenClipboard
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic)
        .type = 1
        .hBmp = GetClipboardData(CF_BITMAP)
    End With

    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, ApiGetClipBmp)
    Call CloseClipboard '关闭剪切板
End Function

Private Function Laplace(W, H, Gray)

'拉普拉斯算子
'---
'0  1  0
'1 -4  1
'0  1  0
'---
'1  1  1
'1 -8  1
'1  1  1
'---
    
    '---
    '锐化模板
    Dim Template
    Dim ptr
    Dim dst
    Dim x, y, n, m, sum, temp
    
    ReDim dst(W - 1, H - 1)
    
    '如果使用的模板中心是负数，那么必须将原图像减去经拉普拉斯变换后的图像。
    '如果使用的模板中心是正数，那么就将原图像加上经过拉普拉斯变换后的图像。
    'Template = Array(0, 1, 0, 1, -4, 1, 0, 1, 0)
    'Template = Array(1, 1, 1, 1, -8, 1, 1, 1, 1)
    Template = Array(0, -1, 0, -1, 4, -1, 0, -1, 0)
    'Template = Array(-1, -1, -1, -1, 8, -1, -1, -1, -1)
    '---
    
    '---
    '锐化滤波
        For x = 0 To W - 1
            For y = 0 To H - 1
                
                '边界处理
                If x = 0 Or x = W - 1 Or y = 0 Or y = H - 1 Then
                    dst(x, y) = Gray(x, y)
                Else
                    
                    If Gray(x, y) = 255 Then
                        x = x
                    End If
                    
                    dst(x, y) = 0
                    sum = 0
                    For m = -1 To 1
                        For n = -1 To 1
                            ptr = Gray(x + m, y + n)
                            sum = sum + ptr * Template(3 * (m + 1) + n + 1)
                        Next
                    Next
                    '-------
                        If sum < 0 Then sum = 0
                        temp = sum + Gray(x, y)
                        If temp > 255 Then temp = 255
                        dst(x, y) = temp
                    '-------
                End If
                
            Next
        Next
    '---
 
    
        For x = 0 To W - 1
            For y = 0 To H - 1
                Gray(x, y) = dst(x, y)
            Next
        Next
'============

End Function
Private Function Resize_1(W, H, reNum, Gray)
        
    Dim x, y, x0, y0, x1, y1, x2, y2, fW, fH
    Dim fy1, fy2, fx1, fx2
    Dim s1, s2, s3, s4
    Dim c1, c2, c3, c4
    
    For x = 0 To W - 1
        x0 = x * (1 / reNum)
        If x0 > W - 1 Then x0 = W - 1
        
        For y = 0 To H - 1

            y0 = y * (1 / reNum)
            If y0 > H - 1 Then y0 = H - 1
            Gray(x, y) = Gray(x0, y0)
 
        Next y
    Next x
    
End Function

Private Function Resize_2(W, H, tW, tH, Gray)
        
    Dim x, y, x0, y0, x1, y1, x2, y2, fW, fH
    Dim fy1, fy2, fx1, fx2
    Dim s1, s2, s3, s4
    Dim c1, c2, c3, c4
    
    fW = W / tW
    fH = H / tH
    
    For x = 0 To W - 1
        x0 = x * fW
        If x0 > W - 1 Then x0 = W - 1
        x1 = CInt(x0 + nOffset)
        If x1 > W - 1 Then x1 = W - 1
        If x > W - 1 Then x2 = x1 Else x2 = x1 + 1
        If x2 > W - 1 Then x2 = W - 1
        fx1 = x1 - x0
        fx2 = 1 - fx1
        
        For y = 0 To H - 1
            If c1 = 0 Then
                y = y
            End If
            
            y0 = y * fH
            If y0 > H - 1 Then y0 = H - 1
            y1 = CInt(y0 + nOffset)
            If y1 > H - 1 Then y1 = H - 1
            If y1 > H - 1 Then y2 = y1 Else y2 = y1 + 1
            If y2 > H - 1 Then y2 = H - 1
            
            fy1 = y1 - y0
            fy2 = 1 - fy1
            
            s1 = fx1 * fy1
            s2 = fx2 * fy1
            s3 = fx2 * fy2
            s4 = fx1 * fy2
            
            c1 = Gray(x1, y1)
            c2 = Gray(x2, y1)
            c3 = Gray(x1, y2)
            c4 = Gray(x2, y2)
            
            temp = c1 * s1 + c2 * s2 + c3 * s4 + c4 * s3
            
            temp = CInt(temp + nOffset)
            
            If temp > 255 Then temp = 255
            If temp < 0 Then temp = 0
            
            Gray(x, y) = temp
 
        Next y
    Next x
    
End Function

Private Function Resize_3(W, H, reNum, Gray)
        
    Dim sW, SH, dW, dH
    Dim B, n, x, y
    Dim i, j
    
    sW = W - 1: SH = H - 1: dW = CInt(W * reNum - 1 + nOffset): dH = CInt(H * reNum - 1 + nOffset)
    
    For i = 0 To dH
        y = i * SH / dH
        n = dH - i * SH Mod dH
        
    Next
        
End Function

Private Function Enforce_BMP_2016() '(Path)
Dim bmp As Bitmap
Dim res As Long
Dim Picture As New StdPicture
Dim Path
Dim myNum
Dim sum
Dim x As Long, y As Long
Dim R, G, B, n, m
Dim i, j, k, t, L, v
Dim temp
Dim W, H, WxH
Dim Flag

'------
Dim SetFlag, GrayFlag, WriteFlag
SetFlag = 0
GrayFlag = 0
'===============
'===============
 

''==============
    Path = "C:\1\1-4.bmp"
    Set Picture = LoadPicture(Path)
    Call GetObject(Picture.Handle, Len(bmp), bmp)
    BMPHandle = Picture.Handle
    W = bmp.bmWidth
    H = bmp.bmHeight
    WxH = W * H
''==============

 
    
'=============
    Dim ImageData() As Byte
    '------
    'bmp.bmWidth须为4的倍数
    ReDim ImageData(0 To 2, 0 To W - 1, 0 To H - 1)
    GetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)
    '------

'=====
    '灰度值处理
    Dim Gray(), tGray, tempGray
    ReDim Gray(W - 1, H - 1)
    For y = 0 To H - 1
        For x = 0 To W - 1
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            Gray(x, y) = Int(R * 0.3 + G * 0.59 + B * 0.11)
        Next x
    Next y

'=====

GrayFlag = 1
'Call Laplace(W, H, Gray)
GetPictureTextFlag = 1
'Resize_1 W, H, 0.5, Gray
Resize_2 W, H, W * 0.9, H * 0.9, Gray


n = 1
If n = 1 Then
    t = 2
    For y = 0 To H - 1
        For x = 0 To W - 1
            '---
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            '---
            
            '=============
            '灰度值处理
            If GrayFlag = 1 Then
                tGray = Gray(x, y) * 1
                If tGray > 255 Then tGray = 255
                If tGray < 0 Then tGray = 0
                R = tGray: G = tGray: B = tGray
            Else
                temp = 0.9
                R = CInt(R * temp + nOffset)
                G = CInt(G * temp + nOffset)
                B = CInt(B * temp + nOffset)
            End If
            '=============
            
            '显示在Excel
            '=====
            If WriteFlag = 1 Then
                If x > 0 And y > 0 And x < W - 1 And y < H - 1 Then
                    Cells(y + 2, x + 1) = ""
                    Cells(y + 2, x + 1).Interior.Color = xlNone
                    If R <> 255 Then
                        Cells(y + 2, x + 1) = R
                        Cells(y + 2, x + 1).Interior.Color = RGB(R, R, R)
                    End If
                End If
            End If
            '=====
            
            '============
            '赋值
                '----
                'R
                ImageData(2, x, y) = R
                'G
                ImageData(1, x, y) = G
                'B
                ImageData(0, x, y) = B
                '---
            '============
        Next x
        t = t + 1
    Next y

End If
'=============
 
    SetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)
    
'=============



        '=============
            Call OpenClipboard(0&) '打开剪切板
            Call EmptyClipboard '清除当前剪切板中的内容
            Call SetClipboardData(CF_BITMAP, BMPHandle) '放置多个不同格式的数据项
            Call CloseClipboard '关闭剪切板
        '=============


'=============
'=============
'========
    Dim SavePath
    SavePath = "C:\1\2.bmp"
    SavePicture ApiGetClipBmp, SavePath
    'SavePicture ApiGetClipBmp, Path
'========
    
    If GetPictureTextFlag = 1 Then
        Debug.Print GetPictureText(SavePath)
    End If
'##############################################################
  
End Function

Private Function test_Enforce_BMP_2016()

Dim Path

Path = "C:\1\1.bmp"

    Dim i, j, k, t, x, y
    
    Dim hDC, hBitmap, W, H
    Dim BMI As BITMAPINFO
    Dim ImageData() As Byte

    
    W = 220
    H = 30


'=====================
    Dim hwnd As Long
    Dim dm As DEVMODE
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    
'    srcDC = CreateDC("DISPLAY", "", "", dm) '创建设备上下文环境
'    BMPHandle = CreateCompatibleBitmap(srcDC, W, H)
'    '----
'    trgDC = CreateCompatibleDC(srcDC)
'    Call SelectObject(trgDC, BMPHandle) '选择一对象到指定的设备上下文环境中
'    Call StretchBlt(trgDC, 0, 0, W, H, srcDC, L, t, mW, mH, SRCCOPY)
'    '----
'    If srcDC <> 0 Then Call ReleaseDC(BMPHandle, srcDC)    '释放设备环境句柄
'    If trgDC <> 0 Then Call DeleteDC(trgDC) '删除内存环境
'
    'SavePic BMPHandle, Path
'=====================


'===============
Dim bmp As Bitmap

'    Call GetObject(BMPHandle, Len(bmp), bmp)
    
'    Debug.Print bmp.bmWidth
    
'===============

    'ReDim ImageData(0 To 2, 0 To W - 1, 0 To H - 1)
    'GetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)


    With BMI.bmHeader
        .bmSize = 40
        .bmWidth = W
        .bmHeight = H
        .bmPlanes = 1
        .bmBitCount = 24
        .bmCompression = 0
    End With
    
'    For i = 0 To 255
'        BMI.bmColors(i).rgbAlpha = 0
'        BMI.bmColors(i).rgbRed = 0
'        BMI.bmColors(i).rgbGreen = 0
'        BMI.bmColors(i).rgbBlue = 255
'    Next i
 
    ReDim ImageData(0 To 2, 0 To W - 1, 0 To H - 1)
    For y = 0 To H - 1
        For x = 0 To W - 1
            ImageData(2, x, y) = 255
            ImageData(1, x, y) = 0
            ImageData(0, x, y) = 0
        Next x
    Next y

'ByVal hDC As Long, _
'lpInfoHeader As BITMAPINFOHEADER, _
'ByVal dwUsage As Long, _
'lpInitBits As Any, _
'lpInitInfo As BITMAPINFO, _
'ByVal wUsage As Long
    
    hDC = GetDC(0)
    hBitmap = CreateDIBitmap(hDC, BMI.bmHeader, CBM_INIT, ImageData(0, 0, 0), BMI, DIB_RGB_COLORS)
    trgDC = CreateCompatibleDC(0)
    SelectObject trgDC, hBitmap
    'Call StretchBlt(trgDC, -W, 0, W, H, hDC, 0, 0, W, H, SRCCOPY)
    'StretchDIBits trgDC, 0, -H, W, H, 0, 0, W, H, ImageData(0, 0), BMI, 0, vbSrcCopy
    ReleaseDC 0, hDC
    DeleteDC trgDC
        
    SavePic hBitmap, Path
    DeleteObject hBitmap
    
End Function

Private Function SetBMI(BMI As BITMAPINFO, W, H)

    With BMI.bmHeader
        .bmSize = 40
        .bmWidth = W
        .bmHeight = H
        .bmPlanes = 1
        .bmBitCount = 24
        .bmCompression = 0
    End With
    
End Function

Private Function Enforce_BMP_2016_test(Path, L, t) '(Path)
Dim bmp As Bitmap
Dim res As Long
Dim Picture As New StdPicture
Dim myNum
Dim sum
Dim x As Long, y As Long
Dim R, G, B, n, m
Dim i, j, k, v
Dim temp
Dim W, H, WxH
Dim Flag
Dim SavePath
'------
Dim SetFlag, GrayFlag, WriteFlag
SetFlag = 0
GrayFlag = 0
'===============
'===============
'===============
'
    Dim hDC As Long
    Dim hwnd As Long
    Dim dm As DEVMODE
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim BMI As BITMAPINFO
    Dim hBitmap As Long
'===============

 

'==============
    Dim mW, mH
    mW = 110
    mH = 15
    W = mW * 1
    H = mH * 1
    srcDC = CreateDC("DISPLAY", "", "", dm) '创建设备上下文环境
    BMPHandle = CreateCompatibleBitmap(srcDC, W, H)
    
    '----
    trgDC = CreateCompatibleDC(srcDC)
    Call SelectObject(trgDC, BMPHandle) '选择一对象到指定的设备上下文环境中
    '抓屏
    Call StretchBlt(trgDC, 0, 0, W, H, srcDC, L, t, mW, mH, SRCCOPY)
    '----
    
    If srcDC <> 0 Then Call ReleaseDC(BMPHandle, srcDC)    '释放设备环境句柄
    If trgDC <> 0 Then Call DeleteDC(trgDC) '删除内存环境
    
SavePic BMPHandle, Path


'==============
    'Path = "C:\1\1-4.bmp"
    Set Picture = LoadPicture(Path)
    BMPHandle = Picture.Handle

    Call GetObject(BMPHandle, Len(bmp), bmp)
    W = bmp.bmWidth
    H = bmp.bmHeight
    WxH = W * H
'==============

'=============
    Dim ImageData() As Byte
    '------
    'bmp.bmWidth须为4的倍数
    ReDim ImageData(0 To 2, 0 To W - 1, 0 To H - 1)
    GetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)
    '------



'=====
    '灰度值处理
    Dim Gray(), tGray, tempGray
    ReDim Gray(W - 1, H - 1)
    For y = 0 To H - 1
        For x = 0 To W - 1
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            Gray(x, y) = Int(R * 0.3 + G * 0.59 + B * 0.11)
        Next x
    Next y


    For y = 0 To H - 1
        For x = 0 To W - 1
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            Gray(x, y) = Int(R * 0.3 + G * 0.59 + B * 0.11)
            If Gray(x, y) >= 200 Then Gray(x, y) = 255
            R = Gray(x, y)
            G = Gray(x, y)
            B = Gray(x, y)
            If x > 0 And y > 0 And x < W - 1 And y < H - 1 Then
                Cells(y + 2, x + 1) = ""
                Cells(y + 2, x + 1).Interior.Color = xlNone
                If R <> 255 Then
                    Cells(y + 2, x + 1) = R
                    Cells(y + 2, x + 1).Interior.Color = RGB(R, G, B)
                End If
            End If
            
            ImageData(2, x, y) = R
            ImageData(1, x, y) = G
            ImageData(0, x, y) = B
            
        Next x
    Next y
    
    SetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)
    
    SavePic BMPHandle, Path
    
End

'Resize_1 W, H, 1, Gray

'GrayFlag = 1
For y = 0 To H - 1
        For x = 0 To W - 1
            '---
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            '---

            '=============
            '灰度值处理
            If GrayFlag = 1 Then
                tGray = Gray(x, y) * 1
                If tGray > 255 Then tGray = 255
                If tGray < 0 Then tGray = 0
                R = tGray: G = tGray: B = tGray
            Else
                temp = 1 '0.9
                R = CInt(R * temp + nOffset)
                G = CInt(G * temp + nOffset)
                B = CInt(B * temp + nOffset)
            End If
            '=============

            '显示在Excel
            '=====
            If WriteFlag = 1 Then
                If x > 0 And y > 0 And x < W - 1 And y < H - 1 Then
                    Cells(y + 2, x + 1) = ""
                    Cells(y + 2, x + 1).Interior.Color = xlNone
                    If R <> 255 Then
                        Cells(y + 2, x + 1) = R
                        Cells(y + 2, x + 1).Interior.Color = RGB(R, G, B)
                    End If
                End If
            End If
            '=====

            ImageData(2, x, y) = R
            ImageData(1, x, y) = G
            ImageData(0, x, y) = B
        Next x
    Next y


mW = W
mH = H
SetBMI BMI, mW, mH



hDC = GetDC(0)
hBitmap = CreateDIBitmap(hDC, BMI.bmHeader, CBM_INIT, ImageData(0, 0, 0), BMI, DIB_RGB_COLORS)
trgDC = CreateCompatibleDC(0)
SelectObject trgDC, hBitmap
ReleaseDC 0, hDC
DeleteDC trgDC
    
SavePic hBitmap, Path
DeleteObject hBitmap


'SetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)

'SavePic BMPHandle, Path

'=====

End

'GrayFlag = 1
'WriteFlag = 1
'Call Laplace(W, H, Gray)
'GetPictureTextFlag = 1
'Resize_1 W, H, 0.5, Gray
'Resize_2 W, H, W, H, Gray


n = 0
If n = 1 Then
    For y = 0 To H - 1
        For x = 0 To W - 1
            '---
            R = ImageData(2, x, y)
            G = ImageData(1, x, y)
            B = ImageData(0, x, y)
            '---

            '=============
            '灰度值处理
            If GrayFlag = 1 Then
                tGray = Gray(x, y) * 1
                If tGray > 255 Then tGray = 255
                If tGray < 0 Then tGray = 0
                R = tGray: G = tGray: B = tGray
            Else
                temp = 1 '0.9
                R = CInt(R * temp + nOffset)
                G = CInt(G * temp + nOffset)
                B = CInt(B * temp + nOffset)
            End If
            '=============

            '显示在Excel
            '=====
            If WriteFlag = 1 Then
                If x > 0 And y > 0 And x < W - 1 And y < H - 1 Then
                    Cells(y + 2, x + 1) = ""
                    Cells(y + 2, x + 1).Interior.Color = xlNone
                    If R <> 255 Then
                        Cells(y + 2, x + 1) = R
                        Cells(y + 2, x + 1).Interior.Color = RGB(R, G, B)
                    End If
                End If
            End If
            '=====

            '============
            '赋值
                '----
                'R
                ImageData(2, x, y) = R
                'G
                ImageData(1, x, y) = G
                'B
                ImageData(0, x, y) = B
                '---
            '============
        Next x
    Next y

End If
'=============

    SetBitmapBits BMPHandle, W * H * 3, ImageData(0, 0, 0)

'=============



        '=============
            Call OpenClipboard(0&) '打开剪切板
            Call EmptyClipboard '清除当前剪切板中的内容
            Call SetClipboardData(CF_BITMAP, BMPHandle) '放置多个不同格式的数据项
            Call CloseClipboard '关闭剪切板
        '=============


'=============
'=============
'========
 
    SavePath = "C:\1\2.bmp"
    SavePath = Path
    SavePicture ApiGetClipBmp, SavePath
    'SavePicture ApiGetClipBmp, Path
'========
    'GetPictureTextFlag = 1
    If GetPictureTextFlag = 1 Then
        Debug.Print GetPictureText(SavePath)
    End If
'##############################################################
  
End Function


Private Function SavePic(BMPHandle, Path)
    '=============
        Call OpenClipboard(0&) '打开剪切板
        Call EmptyClipboard '清除当前剪切板中的内容
        Call SetClipboardData(CF_BITMAP, BMPHandle) '放置多个不同格式的数据项
        Call CloseClipboard '关闭剪切板
    '=============

SavePath = Path
SavePicture ApiGetClipBmp, SavePath
    
End Function
