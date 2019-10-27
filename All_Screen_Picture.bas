Attribute VB_Name = "All_Screen_Picture"
'====================================

Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long
 
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CountClipboardFormats Lib "user32" () As Long
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hDC As Long) As Long

 
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
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
  
Private Sub test_GetPrintScreen()
Dim Path As String

    Path = "F:\AVERY DENNISON\MyCode\New Code_A_Star\Z For RIMSII\1.bmp"
    GetPrintScreen 1, 1, 100, 100, Path
End Sub


Public Function GetPrintScreen(SavePath As String, Left, Top, Width, Height, Optional MultipleW = 1, Optional MultipleH = 1, Optional dwRopFlag As Integer = 1)
    
    Call CaptureScreen(Left, Top, Width, Height, MultipleW, MultipleH, dwRopFlag)
    
    If CountClipboardFormats = 0 Then
        MsgBox "Clipboard is currently empty.", , "Prompt"
        Exit Function
    End If
    
    SavePicture ApiGetClipBmp, SavePath

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
 
'2016-07-28 Star He 屏幕截屏
Private Function CaptureScreen(Left, Top, Width, Height, Optional MultipleW = 1, Optional MultipleH = 1, Optional dwRopFlag As Integer = 1)
    Dim srcDC As Long
    Dim trgDC As Long
    Dim BMPHandle As Long
    Dim dm As DEVMODE
    Dim dwRop As Long
    
    Select Case dwRopFlag
        Case 1
            dwRop = SRCCOPY
        Case 2
            dwRop = NOTSRCCOPY
    End Select
    
    Dim W, H
    W = Width * MultipleW
    H = Height * MultipleH
    
'    If W Mod 4 <> 0 Then
'        MsgBox "请重新设置参数满足宽为4的倍数,现在为：" & W
'        End
'    End If
    
    srcDC = CreateDC("DISPLAY", "", "", dm) '创建设备上下文环境
    trgDC = CreateCompatibleDC(srcDC) '创建与设备相关的内存环境
    BMPHandle = CreateCompatibleBitmap(srcDC, (Width * MultipleW), (Height * MultipleH))
    Call SelectObject(trgDC, BMPHandle) '选择对象
    Call StretchBlt(trgDC, 0, 0, W, H, srcDC, Left, Top, Width, Height, dwRop)   'SRCCOPY,NOTSRCCOPY
    Call OpenClipboard(0&) '打开剪切板
    Call EmptyClipboard '清除当前剪切板中的内容
    Call SetClipboardData(CF_BITMAP, BMPHandle) '放置多个不同格式的数据项
    Call ReleaseDC(BMPHandle, srcDC) '释放设备环境句柄
    Call DeleteDC(trgDC) '删除内存环境
    Call CloseClipboard '关闭剪切板
End Function


