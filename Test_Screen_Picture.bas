Attribute VB_Name = "Test_Screen_Picture"
'2016-08-26
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long
 
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long

Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" _
                                                    (ByVal lpDriverName As String, _
                                                    ByVal lpDeviceName As String, _
                                                    ByVal lpOutput As String, _
                                                    lpInitData As DEVMODE) As Long

'CreateCompatibleDC�����ڴ��н���һ���豸����
'hdcΪҪ�������豸�����ľ��
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

'CreateCompatibleBitmap��������ָ���豸������ص��豸����λͼ
'hdcΪ�豸�������
'nWidthΪָ��λͼ�Ŀ�ȣ���λΪ����
'nHeightΪָ��λͼ�ĸ߶�,��λΪ����
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'SelectObject:��ѡ���һ������ŵ�ָ�����豸������,�¶����滻��ǰ����ͬ���͵Ķ���
'�ڵ�ǰ�豸����ѡ��ͼ�����
'hdcΪ�豸�������
'hObjectΪ��ѡ�����ľ�����ö��������λͼ�ȣ��ұ�����ָ���ĺ�������
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long

'BitBlt����һ��λͼ��һ���豸����������һ��
'��ָ�����豸���������е����ؽ���λ��ת���������͵���һ���豸������
'����λ��ת�����豸������ΪԴ�豸������Ҫ���͵��豸������ΪĿ���豸������Դ��Ŀ�껷�������໥���ݡ�
'hDestDC ΪĿ���豸�������
'xΪĿ���豸���������������Ͻǵ�X�߼�����
'yΪĿ���豸���������������Ͻǵ�Y�߼�����
'nWidthΪԴ��Ŀ�����������߼����
'nHeightΪԴ��Ŀ�����������߼��߶�
'hSrcDCΪԴ�豸�������
'xSrcΪԴ�豸���������������Ͻǵ�X�߼�����
'ySrcΪԴ�豸���������������Ͻǵ�Y�߼�����
'dwRopΪָ����դ�������
Private Declare Function BitBlt Lib "gdi32.dll" _
                                        (ByVal hDestDC As Long, _
                                         ByVal x As Long, _
                                         ByVal y As Long, _
                                         ByVal nWidth As Long, _
                                         ByVal nHeight As Long, _
                                         ByVal hSrcDC As Long, _
                                         ByVal xSrc As Long, _
                                         ByVal ySrc As Long, _
                                         ByVal dwRop As Long) As Long

Private Declare Function StretchBlt Lib "gdi32" _
                                        (ByVal hDC As Long, _
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

'SetPixel:��ָ�����豸����������һ�����ص�RGBֵ
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal crColor As Long) As Long
'SetPixel:��ָ�����豸����������һ�����ص�RGBֵ
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, _
                                                ByVal x As Long, _
                                                ByVal y As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
'DeleteDC:ɾ��ָ�����豸���������ͷ���صĴ�����Դ
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
'DeleteObject:ɾ�������ͷ�������ö����йص�ϵͳ��Դ
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

 
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
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

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32

'  Ternary raster operations
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
'
'=========================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                                                    (Desitination As Any, _
                                                     Source As Any, _
                                                     ByVal Length As Long)
'=========================

'====
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, _
                                                ByVal hBitmap As Long, _
                                                ByVal nStartScan As Long, _
                                                ByVal nNumScans As Long, _
                                                lpBits As Any, _
                                                lpBI As BITMAPINFO, _
                                                ByVal wUsage As Long) As Long

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

Private Sub ByteMe(tempVar)
    
    If tempVar > 255 Then tempVar = 255: Exit Sub
    If tempVar < 150 Then tempVar = 0: Exit Sub
    
End Sub

'2016-09-01
Private Function test_GetBitmatBits()
Dim bmp As Bitmap
Dim res As Long
Dim Picture As New StdPicture
Dim Path
Dim Brightness
Dim myNum
    
    myNum = 90
    Brightness = CSng(myNum) / 100
    Dim i
    Dim bTable(0 To 255)
    Dim tempColor
    For i = 0 To 255
        tempColor = Int(CSng(i) * Brightness)
        ByteMe tempColor
        bTable(i) = tempColor
    Next i
    
    
    Path = "C:\1\1.bmp"
    'Path = "C:\1\3.bmp"
    Set Picture = LoadPicture(Path)
    res = GetObject(Picture.Handle, Len(bmp), bmp)

    
'    Debug.Print "bmBits:" & bmp.bmBits
'    Debug.Print "bmType:" & bmp.bmType
'    Debug.Print "bmBitsPixel:" & bmp.bmBitsPixel
'    Debug.Print "bmWidth:" & bmp.bmWidth
'    Debug.Print "bmHeight:" & bmp.bmHeight

    
    
    Dim ImageData() As Byte
    ReDim ImageData(0 To 2, 0 To bmp.bmWidth - 1, 0 To bmp.bmHeight - 1)
    'ImageData(0,0,0)
    'The first dimension: red(2),green(1),blue(0)
    'The second dimension will be used to address the x coordinates of the image's pixels.
    'The second dimension will be used to address the y coordinates of the image's pixels.
    GetBitmapBits Picture.Handle, bmp.bmWidthBytes * bmp.bmHeight, ImageData(0, 0, 0)
    
'    Debug.Print ImageData(0, 100, 100) 'Blue
'    Debug.Print ImageData(1, 100, 100) 'Green
'    Debug.Print ImageData(2, 100, 100) 'Red

    Dim x As Long, y As Long
    
    For x = 0 To bmp.bmWidth - 1
        For y = 0 To bmp.bmHeight - 1
            'R
            ImageData(2, x, y) = bTable(ImageData(2, x, y))
            'G
            ImageData(1, x, y) = bTable(ImageData(1, x, y))
            'B
            ImageData(0, x, y) = bTable(ImageData(0, x, y))
        Next y
    Next x


           
            '==========
            'Cells(t, x + 1) = tGray
'            If tGray <> 255 Then 'y=8
'                Cells(t, x + 1).Interior.Color = RGB(tGray, tGray, tGray)
'                Cells(t, x + 1) = R
'            Else
'                Cells(t, x + 1) = ""
'            End If
            '==========

'==============================
'            Select Case tGray
'
'                Case 255
'
'                Case 170 To 254
'
'                    'tGray = 255
'
'                Case 0 To 150
'
''                    tGray = 0
'
'                Case Else
''                    k = x + 2
''                    If k < bmp.bmWidth - 1 Then
''                        temp = ImageData(2, k, y)
''                        If temp = 255 Then
''                            tGray = 0
''                        End If
''                    End If
''
''                    k = x - 2
''                    If k < bmp.bmWidth - 1 Then
''                        temp = ImageData(2, k, y)
''                        If temp = 255 Then
''                            tGray = 0
''                        End If
''                    End If
'            End Select
'==============================
            
    SetBitmapBits Picture.Handle, bmp.bmWidthBytes * bmp.bmHeight, ImageData(0, 0, 0)
    
    
    Dim BMPHandle As Long
    BMPHandle = Picture.Handle
    Call OpenClipboard(0&) '�򿪼��а�
    Call EmptyClipboard '�����ǰ���а��е�����
    Call SetClipboardData(CF_BITMAP, BMPHandle) '���ö����ͬ��ʽ��������
    Call CloseClipboard '�رռ��а�
    
    Dim SavePath
    SavePath = "C:\1\2.bmp"
    SavePicture ApiGetClipBmp, SavePath
 
End Function

Private Function ApiGetClipBmp() As IPicture '�Ѽ��а�ͼƬ����ת��ΪͼƬ
 
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
    Call CloseClipboard '�رռ��а�
End Function






'2016-07-28 Star He ��Ļ����
Private Function CaptureScreen_Test(Left, Top, Width, Height)
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim dm As DEVMODE
 
    srcDC = CreateDC("DISPLAY", "", "", dm) '�����豸�����Ļ���
    BMPHandle = CreateCompatibleBitmap(srcDC, Width * 2, Height * 2)
    trgDC = CreateCompatibleDC(srcDC) '�������豸��ص��ڴ滷��
    Call SelectObject(trgDC, BMPHandle) 'ѡ�����
 
    Call StretchBlt(trgDC, 0, 0, Width * 2, Height * 2, srcDC, Left, Top, Width, Height, SRCCOPY)  'SRCCOPY NOTSRCCOPY
    
    Call OpenClipboard(0&) '�򿪼��а�
    Call EmptyClipboard '�����ǰ���а��е�����
    Call SetClipboardData(CF_BITMAP, BMPHandle) '���ö����ͬ��ʽ��������
    Call CloseClipboard '�رռ��а�
    
    Call ReleaseDC(BMPHandle, srcDC) '�ͷ��豸�������
    Call DeleteDC(trgDC) 'ɾ���ڴ滷��
    
End Function


Private Sub test_GetPrintScreen()
Dim Path As String

    Path = "C:\1\1.bmp"
    GetPrintScreen 1, 1, 100, 100, Path
End Sub
 
Private Function GetPrintScreen(Left, Top, Width, Height, SavePath As String)
    
    Call CaptureScreen_Test(Left, Top, Width, Height)
    
    If CountClipboardFormats = 0 Then
        MsgBox "Clipboard is currently empty.", , "Prompt"
        Exit Function
    End If
    
    SavePicture ApiGetClipBmp, SavePath

End Function

'
'12-30 ��Ļ����
Private Function CaptureScreen(Left As Integer, Top As Integer, Width As Integer, Height As Integer)
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim dm As DEVMODE

    srcDC = CreateDC("DISPLAY", "", "", dm) '�����豸�����Ļ���
    trgDC = CreateCompatibleDC(srcDC) '�������豸��ص��ڴ滷��
    BMPHandle = CreateCompatibleBitmap(srcDC, Width, Height)
    Call SelectObject(trgDC, BMPHandle) 'ѡ�����
    Call BitBlt(trgDC, 0, 0, Width, Height, srcDC, Left, Top, SRCCOPY) 'SRCCOPY) 'λͼ����
    Call OpenClipboard(0&) '�򿪼��а�
    Call EmptyClipboard '�����ǰ���а��е�����
    Call SetClipboardData(CF_BITMAP, BMPHandle) '���ö����ͬ��ʽ��������
    Call ReleaseDC(BMPHandle, srcDC) '�ͷ��豸�������
    Call DeleteDC(trgDC) 'ɾ���ڴ滷��
    Call CloseClipboard '�رռ��а�
End Function



'VarPtr����������Ϊ��ȡ�����ĵ�ַ
Private Function test_CopyMemory()
'    Dim long1 As Long
'    Dim long2 As Long
'
'    long1 = 10
'
'    MsgBox LenB(long1)
'
'    CopyMemory1 long2, long1, 4
'    MsgBox long2
'
'    Dim long3 As Long
'    CopyMemory2 VarPtr(long3), VarPtr(long1), 4
'
'    MsgBox long3

Dim Arr(1 To 3) As Integer
For i = 1 To 3
    Arr(i) = i * 10
Next

'Dim arraddr As Long
'arraddr = VarPtrArray(arr)

Dim SafeArrayldPoint As Long

CopyMemory SafeArrayldPoint, ByVal VarPtrArray(Arr), 4

Dim dims As Integer
CopyMemory dims, ByVal SafeArrayldPoint, 2

Dim elements As Long
CopyMemory elemnets, ByVal SafeArrayldPoint + 4, 4

Dim eCount As Long
CopyMemory eCount, ByVal SafeArrayldPoint + 16, 4

Dim lBd As Long
CopyMemory lBd, ByVal SafeArrayldPoint + 20, 4

'��ȡ�����ֵ
Dim arraddr As Long
CopyMemory arraddr, ByVal SafeArrayldPoint + 12, 4

Dim arr1 As Integer, arr2 As Integer, arr3 As Integer

CopyMemory arr1, ByVal arraddr, 2
CopyMemory arr2, ByVal arraddr + 2, 2
CopyMemory arr3, ByVal arraddr + 4, 2

Debug.Print arr1
Debug.Print arr2
Debug.Print arr3

CopyMemory ByVal arraddr, 13, 2
CopyMemory ByVal arraddr + 2, 28, 2

Debug.Print Arr(1)
Debug.Print Arr(2)
Debug.Print Arr(3)



'arraddr = VarPtrArray(arr)
'
'Dim arr1 As Integer, arr2 As Integer, arr3 As Integer
'
'CopyMemory arr1, arraddr, LenB(arr1)
'CopyMemory arr2, arraddr + 2, 2
'CopyMemory arr3, arraddr + 4, 2
'
'Debug.Print arr1
'Debug.Print arr2
'Debug.Print arr3

End Function


'Star He 2016-07-27
Private Function Get_OrderNO() As String
 
Dim hwnd As Long
Dim AllHwnd() As String

Dim i As Integer
Dim j As Integer
Dim t As Long
Dim temp As String
Dim StrData As String
Dim Myhwnd As Long
Dim MyRect As RECT
    
        Call Init
        hwnd = setHwnd
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        'ȥ���������Ч�ַ�
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        'ת��������
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------
    

Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim Path As String
Dim PicStr As String


        Path = "C:\1\1.bmp"
        If Dir("C:\1\", vbDirectory) = "" Then
            MkDir "C:\1"
        End If
    '--------------
        Myhwnd = CLng(AllHwnd(27))
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '��ȡ��ӦRectangle����,Order No.�ľ�������
        '===================
            
            t = 1
            
            If t = 1 Then
                x0 = MyRect.Left + 356 '+ 177 '132
                y0 = MyRect.Top + 38 ' + 41 '41
                PicStr = Get_Picture_Str(Path, x0, y0, 46, 14, 2, 2)
                Debug.Print PicStr
            Else
                x0 = MyRect.Left + 179 '132
                y0 = MyRect.Top + 42 '41
                PicStr = Get_Picture_Str(Path, x0, y0, 46, 14, 2, 2) '����94,��15��W�Ŵ�����4,H�Ŵ�����6
                Debug.Print PicStr
            End If
            
            End
 
'
            
'            Do
'                DoEvents
'                PicStr = Get_Picture_Str(Path, x0, y0, 49, 16, 1, 1) '����94,��15��W�Ŵ�����4,H�Ŵ�����6
'                PicStr = VBA.Replace(PicStr, " ", "")
'                Debug.Print PicStr
'                Exit Do
'            Loop Until (VBA.Len(PicStr) = 7)
            
             
        '===================
        
        '===================
            x0 = MyRect.Left + 177 '132
            y0 = MyRect.Top + 41 '41
            'x1 = MyRect.Right
            'y1 = MyRect.Bottom
            'Debug.Print x0, y0
            'Debug.Print x1, y1
            'GetCursorPos XY
            'Debug.Print XY.x, XY.y
            'End
'            For i = 1 To 6
'                For j = 1 To 10
'                    DoEvents
'                    PicStr = Get_Picture_Str(Path, x0, y0, 49, 16, i, j)
'                    Debug.Print i & "," & j & ": " & PicStr
'                    If PicStr = "Image is not readable" Then
'                        Exit For
'                    End If
'                Next j
'            Next i
'
            
            Do
                DoEvents
                PicStr = Get_Picture_Str(Path, x0, y0, 49, 16, 1, 1) '����94,��15��W�Ŵ�����4,H�Ŵ�����6
                PicStr = VBA.Replace(PicStr, " ", "")
                Debug.Print PicStr
                Exit Do
            Loop Until (VBA.Len(PicStr) = 7)
            
            End
        '===================
        temp = PicStr
        temp = VBA.Replace(temp, VBA.Chr(0), "")
        If temp = "Image is not readable" Or VBA.Trim(temp) = "" Then
            PicStr = ""
        End If
        
        'PicStr = VBA.Replace(PicStr, VBA.Chr(0), "")
        'PicStr = VBA.Replace(PicStr, " ", "")
        Get_OrderNO = PicStr
        
        Debug.Print PicStr
    '--------------
End Function

'2016-07-28
Private Function Popup_Search_test_old() '(AddressNO As String)
Dim AddressNO_temp As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As Integer

Dim hwnd As Long
Dim Myhwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Data() As String
Dim MyRect As RECT
Dim Title As String
    
    If GetESC Then SetESC
    
    Title = "Popup Search"
    hwnd = FindWindow(vbNullString, Title)
    t = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Debug.Print "��ESCֹͣ����,����Ϊû���ҵ�:" & Title
            End
        End If
        hwnd = FindWindow(vbNullString, Title)
        Delay 100
        t = t + 1
        If t > 15 Then 'ѭ��1.5��
            Debug.Print "û�д�:" & Title
            Exit Function
        End If
    Loop
 
 
    Set ws = CreateObject("WSCRIPT.SHELL")
    ws.AppActivate Title
    Delay 300
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        'ȥ���������Ч�ַ�
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        'ת��������
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------
    

Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim Path As String
Dim Flag As Integer
Dim PicStr As String
Dim myrow As Integer
Dim Times As Integer
Dim DownFlag As Integer
        
        'Path = "F:\AVERY DENNISON\MyCode\New Code_A_Star\Z For RIMSII\1.bmp"
        Path = "C:\1\1.bmp"
        If Dir("C:\1\", vbDirectory) = "" Then
            MkDir "C:\1"
        End If
    '--------------
        Myhwnd = CLng(AllHwnd(3)) 'Master Customer
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '��ȡItem ID��ӦRectangle����
        x0 = MyRect.Left + 32 '32 30-32
        y0 = MyRect.Top + 17 '17 20,16
        PicStr = Get_Picture_Str(Path, x0, y0, 88, 143, 2, 2, 2) 'x0,y0:��ʼλ�����꣬���Ŵ�����2 ��Ŵ�����2 ,����dwRopFlagĬ��Ϊ1
        temp = PicStr
        
        Debug.Print temp
        
        End
        x0 = MyRect.Left + 43 '32 30-32
        y0 = MyRect.Top + 17 '17 20,16

        '-------
            'Debug.Print Round((x1 - x0) / 4) - 45 88
            'Debug.Print (y1 - y0 - 10) 143
            
            Dim AddressNO As String
            AddressNO = "6200/1" '4596/9
            PicStr = Get_Picture_Str(Path, x0, y0, 70, 160, 2, 2, 2) 'x0,y0:��ʼλ�����꣬���Ŵ�����2 ��Ŵ�����2 ,����dwRopFlagĬ��Ϊ1
            temp = PicStr
            
            Debug.Print temp
            End
        '-------
            
        DownFlag = 0
        Do
            DoEvents
            
            '��ȡ��Ϣ
            PicStr = Get_Picture_Str(Path, x0, y0, Round((x1 - x0) / 4) - 45, (y1 - y0 - 10))
            temp = PicStr
            
            Debug.Print temp
            
            If temp = "Image is not readable" Or VBA.Trim(temp) = "" Then
                MsgBox "û���ҵ���ȷ��Address NO.!"
                End
            End If
            
            If InStr(temp, VBA.Chr(10)) Then
                Erase Data
                Data = VBA.Split(temp, VBA.Chr(10))
                Flag = 0
                For i = 0 To UBound(Data)
                    temp = VBA.Replace(VBA.Trim(Data(i)), " ", "")
                    If InStr(temp, AddressNO) Then
                        Flag = 1
                        Exit For
                    End If
                Next i
                
                If Flag = 1 Then '�ҵ���Ӧ����
                    'ѡ�ж�Ӧ����
                    '-----
                        myrow = Get_Blue_Row(x0, y0, x1, y1)
                        Times = myrow - 1
                        t = i - Times
                        If t < 0 Then
                            For j = 1 To VBA.Abs(t)
                                SendMyData "Up"
                            Next j
                        Else
                            For j = 1 To t
                               SendMyData "Down"
                            Next j
                        End If
                    '-----
                    '��ʱ
                    Delay 200
                    'Select
                    '----
                        Myhwnd = CLng(AllHwnd(2)) 'Select
                        PostMessage Myhwnd, BM_CLICK, 0, 0 '���Select��ť.
                    '----
                    Exit Do '����ѭ��
                Else
                    '---
                    If DownFlag = 0 Then
                        SendMyData "Down" '��ʱ��ѡ���е���ɫ����Ӱ����Ļ��ȡ�Ľ��������Ҫ�����ƶ�һ�Σ����ų�Ӱ��
                        DownFlag = 1
                    Else
                        If UBound(Data) = 6 Then
                            SendMyData "PgDn"
                            Delay 500
                            DownFlag = 0 '��ҳ��������µ�ҳ��ʵ������һ��
                        Else
                            Erase Data
                            If InStr(AddressNO, "/") Then
                                Data = VBA.Split(AddressNO, "/")
                                AddressNO_temp = "/" & Data(1)
                            End If
                            Erase Data
                            temp = PicStr
                            Data = VBA.Split(temp, VBA.Chr(10))
                            Flag = 0
                            For i = 0 To UBound(Data)
                                temp = VBA.Replace(VBA.Trim(Data(i)), " ", "")
                                If InStr(temp, AddressNO_temp) Then
                                    Flag = 1
                                    Exit For
                                End If
                            Next i
                            
                            If Flag = 0 Then
                                MsgBox "û���ҵ�Address NO.:" & AddressNO & "!" & vbCrLf & PicStr
                                End
                            Else
                                AddressNO = AddressNO_temp
                            End If
                        End If 'If UBound(Data) = 6 Then
                    End If 'If DownFlag = 0 Then
                    '---
                End If
            End If
        Loop
    '--------------
    
    '--------------
End Function
'2016-07-28
Private Function Popup_Search_test() '(AddressNO As String)
Dim AddressNO_temp As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As Integer

Dim hwnd As Long
Dim Myhwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Data() As String
Dim MyRect As RECT
Dim Title As String
    
    If GetESC Then SetESC
    
    Title = "Popup Search"
    hwnd = FindWindow(vbNullString, Title)
    t = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Debug.Print "��ESCֹͣ����,����Ϊû���ҵ�:" & Title
            End
        End If
        hwnd = FindWindow(vbNullString, Title)
        Delay 100
        t = t + 1
        If t > 15 Then 'ѭ��1.5��
            Debug.Print "û�д�:" & Title
            Exit Function
        End If
    Loop
 
 
    Set ws = CreateObject("WSCRIPT.SHELL")
    ws.AppActivate Title
    Delay 300
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        'ȥ���������Ч�ַ�
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        'ת��������
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------
    

Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim Path As String
Dim Flag As Integer
Dim PicStr As String
Dim myrow As Integer
Dim TargetRow As Integer
Dim Times As Integer
Dim DownFlag As Integer
Dim MyLastRow As Integer
Dim OneNote As Object

        'Path = "F:\AVERY DENNISON\MyCode\New Code_A_Star\Z For RIMSII\1.bmp"
        Path = "C:\1\1.bmp"
        If Dir("C:\1\", vbDirectory) = "" Then
            MkDir "C:\1"
        End If
    '--------------
        Myhwnd = CLng(AllHwnd(3)) 'Master Customer
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '��ȡItem ID��ӦRectangle����
        
        'For Use
        '-----------
            x0 = MyRect.Left + 43 '32 30-32
            y0 = MyRect.Top + 15 '17 20,16
            x1 = MyRect.Right
            y1 = MyRect.Bottom
        '-----------
        'MyRow = Get_Blue_Row(x0, y0, x1, y1)
        MyLastRow = Get_Last_Row(x0, y0, x1, y1)
        
        'End
        
        Dim MyRange As ScreenRange
        Dim Tdata() As String

'        '-------
            Dim AddressNO As String
            AddressNO = "4596/11" ' "6200/1" ' "1434/12" '4596/9 6200/1
'        '--------

        'stime = Timer
        
        Call Kill_OneNote
        Set OneNote = CreateObject("OneNote.Application")
        
        Erase Data
        Erase Tdata
        t = 0
        i = 1
        Do
            DoEvents
            If GetESC Then
                Call SetMyCursor(0)
                Debug.Print "All End"
                End
            End If
            
            MyRange = GetRange(i, x0, y0)
            PicStr = Get_Picture_Str(Path, MyRange.Left, MyRange.Top, MyRange.Width, MyRange.Height, 2, 2, 1, OneNote)
            temp = PicStr
            temp = VBA.Replace(temp, " ", "")
            Debug.Print "Row:" & i & " " & temp
            If AddressNO = temp Then
                TargetRow = (i - 1)
                'MyRow = Get_Blue_Row(x0, y0, x1, y1)
                'Debug.Print MyRow
                'Debug.Print TargetRow
                Debug.Print temp
                Exit Do
            End If
                                   
            If i = 7 Then
                SendMyData "PgDn"
                Delay 500
                i = Get_Blue_Row(x0, y0, x1, y1) - 1
                MyLastRow = Get_Last_Row(x0, y0, x1, y1)
            Else
                SendMyData "Down"
            End If
            
            If i = MyLastRow Then
                Exit Do
            End If
                        
            i = i + 1
        Loop
 
        'Call Select_Row(TargetRow, MyRow)
        'etime = Timer
        'Debug.Print "Time:" & etime - stime
        
        End



   
'
'        'New
'        '-------
            x0 = MyRect.Left + 43 '32 30-32
            y0 = MyRect.Top + 15 '17 20,16
            x1 = MyRect.Right
            y1 = MyRect.Bottom

            myrow = Get_Blue_Row(x0, y0, x1, y1)

            MyRange = GetRange(myrow, x0, y0)
            PicStr = Get_Picture_Str(Path, MyRange.Left, MyRange.Top, MyRange.Width, MyRange.Height + 23, 2, 2)
            'GetPrintScreen Path, MyRange.Left, MyRange.Top, MyRange.Width, MyRange.Height
            'PicStr = Get_Picture_Str(Path, x0, y0, 77, 170, , , 1)
            temp = PicStr
            Debug.Print "��" & myrow & "��" & vbCrLf & temp
            Debug.Print "---------------------------"
             
'            PicStr = Get_Picture_Str(Path, x0, y0, 77, 165)
'            Debug.Print PicStr
'            End
            
'            For i = 1 To 7
'                MyRow = i
'
'                Dim MyRange As ScreenRange
'
'                MyRange = GetRange(MyRow, x0, y0)
'                PicStr = Get_Picture_Str(Path, MyRange.Left, MyRange.Top, MyRange.Width, MyRange.Height * i, 2, 2)
'                'GetPrintScreen Path, MyRange.Left, MyRange.Top, MyRange.Width, MyRange.Height
'                'PicStr = Get_Picture_Str(Path, x0, y0, 77, 170, , , 1)
'                temp = PicStr
'                Debug.Print "��" & MyRow & "��" & vbCrLf & temp
'                Debug.Print "---------------------------"
'            Next
            
            
            
            
 
    
    
    
            End
'        '-------
'        'Old
'        '=====================
'            x0 = MyRect.Left + 32 '32 30-32
'            y0 = MyRect.Top + 17 '17 20,16
'            x1 = MyRect.Right
'            y1 = MyRect.Bottom
'
'            MyRow = Get_Blue_Row(x0, y0, x1, y1)
'
'            PicStr = Get_Picture_Str(Path, x0, y0, 88, 143, 2, 2, 1) 'x0,y0:��ʼλ�����꣬���Ŵ�����2 ��Ŵ�����2 ,����dwRopFlagĬ��Ϊ1
'            temp = PicStr
'            Debug.Print temp
'            End
'        '===================
        
        DownFlag = 0
        Do
            DoEvents
            
            '��ȡ��Ϣ
            PicStr = Get_Picture_Str(Path, x0, y0, 77, 165)
            temp = PicStr
            
            Debug.Print temp
            
            If temp = "Image is not readable" Or VBA.Trim(temp) = "" Then
                MsgBox "û���ҵ���ȷ��Address NO.!"
                End
            End If
            
            If InStr(temp, VBA.Chr(10)) Then
                Erase Data
                Data = VBA.Split(temp, VBA.Chr(10))
                Flag = 0
                For i = 0 To UBound(Data)
                    temp = VBA.Replace(VBA.Trim(Data(i)), " ", "")
                    If InStr(temp, AddressNO) Then
                        Flag = 1
                        Exit For
                    End If
                Next i
                
                If Flag = 1 Then '�ҵ���Ӧ����
                    'ѡ�ж�Ӧ����
                    '-----
                        myrow = Get_Blue_Row(x0, y0, x1, y1)
                        Times = myrow - 1
                        t = i - Times
                        If t < 0 Then
                            For j = 1 To VBA.Abs(t)
                                SendMyData "Up"
                            Next j
                        Else
                            For j = 1 To t
                               SendMyData "Down"
                            Next j
                        End If
                    '-----
                    '��ʱ
                    Delay 200
                    'Select
                    '----
                        Myhwnd = CLng(AllHwnd(2)) 'Select
                        PostMessage Myhwnd, BM_CLICK, 0, 0 '���Select��ť.
                    '----
                    Exit Do '����ѭ��
                Else
                    '---
                    If DownFlag = 0 Then
                        SendMyData "Down" '��ʱ��ѡ���е���ɫ����Ӱ����Ļ��ȡ�Ľ��������Ҫ�����ƶ�һ�Σ����ų�Ӱ��
                        DownFlag = 1
                    Else
                        If UBound(Data) = 6 Then
                            SendMyData "PgDn"
                            Delay 500
                            DownFlag = 0 '��ҳ��������µ�ҳ��ʵ������һ��
                        Else
                            Erase Data
                            If InStr(AddressNO, "/") Then
                                Data = VBA.Split(AddressNO, "/")
                                AddressNO_temp = "/" & Data(1)
                            End If
                            Erase Data
                            temp = PicStr
                            Data = VBA.Split(temp, VBA.Chr(10))
                            Flag = 0
                            For i = 0 To UBound(Data)
                                temp = VBA.Replace(VBA.Trim(Data(i)), " ", "")
                                If InStr(temp, AddressNO_temp) Then
                                    Flag = 1
                                    Exit For
                                End If
                            Next i
                            
                            If Flag = 0 Then
                                MsgBox "û���ҵ�Address NO.:" & AddressNO & "!" & vbCrLf & PicStr
                                End
                            Else
                                AddressNO = AddressNO_temp
                            End If
                        End If 'If UBound(Data) = 6 Then
                    End If 'If DownFlag = 0 Then
                    '---
                End If
            End If
        Loop
    '--------------
    
    '--------------
End Function

'Private Function Get_Last_Row(x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer) As Integer
'Dim hdc As Long
'Dim MyColorNum As Long
'Dim i As Integer
'Dim j As Integer
'Dim x As Integer
'Dim y As Integer
'Dim temp As String
'Dim Data
'
'
'    hdc = SetHdc
'
'    Data = Array(10, 30, 48, 67, 86, 106, 125) 'ÿ����ֵ
'
'    For i = 0 To UBound(Data)
'
'        y = y0 + Data(i)
'
'        DoEvents
'
'        MyColorNum = GetPixel(hdc, x0, y)
'        If MyColorNum = 16750899 Then '��ɫ��
'            MyRow = GetMyRow(y0, y) '�ڵڼ���
'            If InStr(temp, MyRow) = 0 Then
'                For x = x0 + 10 To x0 + 20
'                    MyColorNum = GetPixel(hdc, x, y)
'                    If MyColorNum = 16777215 Then ' MyColorNum = 16777215 ��ɫ������Ϊ��ɫ
'                        temp = temp & MyRow
'                        'Debug.Print "test:" & MyRow
'                        Exit For
'                    End If
'                Next
'
'            End If
'        Else
'            MyRow = GetMyRow(y0, y) '�ڵڼ���
'            If InStr(temp, MyRow) = 0 Then
'                For x = x0 + 10 To x0 + 20
'                    MyColorNum = GetPixel(hdc, x, y)
'                    If MyColorNum = 0 Then ' MyColorNum = 0 ����Ϊ��ɫ
'                        temp = temp & MyRow
'                        'Debug.Print "test:" & MyRow
'                        Exit For
'                    End If
'                Next
'            End If
'        End If
'
'    Next
'
'    Get_Last_Row = CInt(VBA.Right(temp, 1))
'
''    temp = ""
''
''    Dim MidNum As Integer
''
''    MidNum = 10
''
''    For y = y0 + 10 To y1 Step MidNum
''
''        Debug.Print y
''
''        '----------
''            If y > (y0 + 134) Then '������Χ������ѭ��
''                Exit For
''            End If
''        '----------
''
''        If y > y0 + 1 Then
''
''            Call SetCursorPos(x0 + 10, y)
''            Debug.Print GetMyRow(y0, y)
''
''        End If
''
''        y = y + MidNum
''
''        Sleep 500
''
''    Next
'
'    Exit Function
'
'    For y = y0 To y1 Step 9
'
'        '----------
'            If y > (y0 + 134) Then '������Χ������ѭ��
'                Exit For
'            End If
'        '----------
'        DoEvents
'
'        'Call SetCursorPos(x0, y)
'        MyColorNum = GetPixel(hdc, x0, y)
'
'        'Debug.Print y
'
'        If MyColorNum = 16750899 Then '��ɫ��
'            MyRow = GetMyRow(y0, y) '�ڵڼ���
'            'Debug.Print MyRow
'            If InStr(temp, MyRow) = 0 Then
'                For x = x0 + 10 To x0 + 20
'                    'Call SetCursorPos(x, y)
'                    MyColorNum = GetPixel(hdc, x, y)
'                    If MyColorNum = 16777215 Then ' MyColorNum = 16777215 ��ɫ������Ϊ��ɫ
'                        temp = temp & MyRow
'                        'Debug.Print "test:" & MyRow
''                        If MyRow < 7 Then
''                            If GetMyRow(y0, y + 9) = MyRow Then '�����һ���жϻ��Ǵ��У���������һ���ж�
''                                y = y + 9
''                            End If
''                        End If
'                        Exit For
'                    End If
'                Next
'
'            End If
'        Else
'            If y > y0 Then
'                MyRow = GetMyRow(y0, y) '�ڵڼ���
'                'Debug.Print MyRow
'                If InStr(temp, MyRow) = 0 Then
'                    For x = x0 + 10 To x0 + 20
'                        'Call SetCursorPos(x, y)
'                        MyColorNum = GetPixel(hdc, x, y)
'                        If MyColorNum = 0 Then ' MyColorNum = 0 ����Ϊ��ɫ
'                            temp = temp & MyRow
''                            If MyRow < 7 Then
''                                If GetMyRow(y0, y + 9) = MyRow Then '�����һ���жϻ��Ǵ��У���������һ���ж�
''                                    y = y + 9
''                                End If
''                            End If
'                            'Debug.Print "test:" & MyRow
'                            Exit For
'                        End If
'                    Next
'                End If
'            End If
'        End If
'
'        If VBA.Len(temp) > 1 Then
'            If MyRow - VBA.Len(temp) > 1 Then '�ж��Ƿ�հ׶��У�����ǣ�������ѭ��������ʱ�䣬����ÿ�ζ��ж�7������
'                Exit For
'            End If
'        End If
'    Next
'
'    Get_Last_Row = CInt(VBA.Right(temp, 1))
'
'
'End Function
