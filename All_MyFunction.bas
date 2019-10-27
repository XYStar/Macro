Attribute VB_Name = "All_MyFunction"
'2014-12-16
'2014-12-23
'2014-12-24

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type

Dim XY As POINTAPI

'该函数检取光标的位置，以屏幕坐标表示。
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long

Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long
Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
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

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'  Device Parameters for GetDeviceCaps()
Private Const DRIVERVERSION = 0      '  Device driver version
Private Const TECHNOLOGY = 2         '  Device classification
Private Const HORZSIZE = 4           '  Horizontal size in millimeters
Private Const VERTSIZE = 6           '  Vertical size in millimeters
Private Const HORZRES = 8            '  Horizontal width in pixels
Private Const VERTRES = 10           '  Vertical width in pixels
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PLANES = 14            '  Number of planes
Private Const NUMBRUSHES = 16        '  Number of brushes the device has
Private Const NUMPENS = 18           '  Number of pens the device has
Private Const NUMMARKERS = 20        '  Number of markers the device has
Private Const NUMFONTS = 22          '  Number of fonts the device has
Private Const NUMCOLORS = 24         '  Number of colors the device supports
Private Const PDEVICESIZE = 26       '  Size required for device descriptor
Private Const CURVECAPS = 28         '  Curve capabilities
Private Const LINECAPS = 30          '  Line capabilities
Private Const POLYGONALCAPS = 32     '  Polygonal capabilities
Private Const TEXTCAPS = 34          '  Text capabilities
Private Const CLIPCAPS = 36          '  Clipping capabilities
Private Const RASTERCAPS = 38        '  Bitblt capabilities
Private Const ASPECTX = 40           '  Length of the X leg
Private Const ASPECTY = 42           '  Length of the Y leg
Private Const ASPECTXY = 44          '  Length of the hypotenuse

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const SIZEPALETTE = 104      '  Number of entries in physical palette
Private Const NUMRESERVED = 106      '  Number of reserved entries in palette
Private Const COLORRES = 108         '  Actual color resolution

 
Private Const OCR_NORMAL = 32512
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


'===============================================================
 
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const BM_CLICK = &HF5
Private Const WM_ACTIVATE = &H6
Private Const WA_ACTIVE = 1
Private Const WM_CLOSE = &H10
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_LEFTUP As Long = &H4
Private Const WM_SETTEXT As Long = &HC
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
 
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long


Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105

Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102


' Virtual Keys, Standard Set
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_CANCEL = &H3
Private Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
Private Const VK_BACK = &H8
Private Const VK_TAB = &H9
Private Const VK_CLEAR = &HC
Private Const VK_ENTER = &HD 'enter
Private Const VK_SHIFT = &H10 'Shift
Private Const VK_CTRL = &H11 'Ctrl
Private Const VK_ALT = &H12 'Alt
Private Const VK_PAUSE = &H13
Private Const VK_CAPITAL = &H14
Private Const VK_ESCAPE = &H1B
Private Const VK_SPACE = &H20
Private Const VK_PRIOR = &H21
Private Const VK_NEXT = &H22
Private Const VK_END = &H23
Private Const VK_HOME = &H24
Private Const VK_LEFT = &H25
Private Const VK_UP = &H26
Private Const VK_RIGHT = &H27
Private Const VK_DOWN = &H28
Private Const VK_SELECT = &H29
Private Const VK_PRINT = &H2A
Private Const VK_EXECUTE = &H2B
Private Const VK_SNAPSHOT = &H2C
Private Const VK_INSERT = &H2D
Private Const VK_DELETE = &H2E
Private Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_MULTIPLY = &H6A
Private Const VK_ADD = &H6B
Private Const VK_SEPARATOR = &H6C
Private Const VK_SUBTRACT = &H6D
Private Const VK_DECIMAL = &H6E
Private Const VK_DIVIDE = &H6F
Private Const VK_F1 = &H70
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
 

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Const MOUSEEVENTF_WHEEL = &H800 '2016-09-12
'==========================
 
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2

Dim AppData As String
Private SetHdc As Long
Dim xScrn As Integer
Dim yScrn As Integer
Dim MyColorNum As String

Public Type ScreenRange
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

Public setHwnd As Long
Public OneNote As Object
Public ws As Object
 
Public Function Delay(ms As Long)
 
    Call Sleep(ms / 2)
    Call Cursor_Change
    DoEvents
    
    Call Sleep(ms / 2)
    Call Cursor_Change
    DoEvents
    
End Function
''###################################################################################################################
''###################################################################################################################
''###################################################################################################################
 
Public Sub SendMyData(Data As String)
Dim i As Integer
Dim hwnd As Long
Dim MyData As String
Dim AscData As String
Dim Flag As Integer
Dim State As Integer
     
    'Init
    
    Debug.Print Data
 
    If GetESC Then
        Debug.Print "All End"
        End
    End If

    Select Case Data
        Case "^a" '全选
            Keybd_Event_Combine VK_CTRL, Asc("A")
            
        Case "^c" '复制
            Keybd_Event_Combine VK_CTRL, Asc("C")
            
        Case "^v" '粘贴
            Keybd_Event_Combine VK_CTRL, Asc("V")
            
        Case "^s" '保存
            Keybd_Event_Combine VK_CTRL, Asc("S")
            
        Case "^n" '
            Keybd_Event_Combine VK_CTRL, Asc("N")
                        
        Case "%f" 'Oracle 文件
            Keybd_Event_Combine VK_ALT, Asc("F")
            
        Case "%e" 'Oracle 编辑
            Keybd_Event_Combine VK_ALT, Asc("E")
            
        Case "%v" 'Oracle 查看
            Keybd_Event_Combine VK_ALT, Asc("V")
            
        Case "%l" 'Oracle 文件夹
            Keybd_Event_Combine VK_ALT, Asc("L")
            
        Case "%t" 'Oracle 工具
            Keybd_Event_Combine VK_ALT, Asc("T")
            
        Case "%o" 'Oracle 确认
            Keybd_Event_Combine VK_ALT, Asc("O")
            
        Case "%m" 'Oracle Submit
            Keybd_Event_Combine VK_ALT, Asc("M")
        
        Case "%d" 'Oracle 打开Zoom
            Keybd_Event_Combine VK_ALT, Asc("D")
        
        Case "%u" 'Zoom Use Double Key Entry
            Keybd_Event_Combine VK_ALT, Asc("U")
            
        Case "%s" 'Zoom Close
            Keybd_Event_Combine VK_ALT, Asc("S")
         
        Case "%y" 'Zoom Copy
            Keybd_Event_Combine VK_ALT, Asc("Y")
            
        Case "%a"
            Keybd_Event_Combine VK_ALT, Asc("A")
         
        Case "^F11"
            Keybd_Event_Combine VK_CTRL, VK_F11
        
        Case "{Space}"
            Keybd_Event_Data VK_SPACE
        
        Case "Tab"
            Keybd_Event_Data VK_TAB

        Case "+Tab"
            Keybd_Event_Combine VK_SHIFT, VK_TAB
            
        Case "{+TAB}"
            Keybd_Event_Combine VK_SHIFT, VK_TAB
            
        Case "Up"
            Keybd_Event_Data VK_UP
         
        Case "Down"
            Keybd_Event_Data VK_DOWN
            
        Case "Right"
            Keybd_Event_Data VK_RIGHT
           
        Case "Left"
            Keybd_Event_Data VK_LEFT

        Case "PgUp"
            Keybd_Event_Data VK_PRIOR
        
        Case "PgDn"
            Keybd_Event_Data VK_NEXT
            
        Case "{F1}"
            Keybd_Event_Data VK_F1
            
        Case "{F2}"
            Keybd_Event_Data VK_F2
             
        Case "{F3}"
            Keybd_Event_Data VK_F3
              
        Case "{F4}"
            Keybd_Event_Data VK_F4
            
        Case "{F5}"
            Keybd_Event_Data VK_F5
            
        Case "{F6}"
            Keybd_Event_Data VK_F6
             
        Case "{F7}"
            Keybd_Event_Data VK_F7
              
        Case "{F8}"
            Keybd_Event_Data VK_F8
            
        Case "{F9}"
            Keybd_Event_Data VK_F9
            
        Case "{F10}"
            Keybd_Event_Data VK_F10
              
        Case "{F11}"
            Keybd_Event_Data VK_F11
            
        Case "{F12}"
            Keybd_Event_Data VK_F12
            
        Case "Enter"
            Keybd_Event_Data VK_ENTER
            
        Case "Back"
            Keybd_Event_Data VK_BACK
            
        Case Else
            Flag = 0
            
            '--------------2015-08-20
            Dim temp As String
            If VBA.Len(Data) = 2 Then
                If VBA.Left(Data, 1) = "%" Then
                    temp = VBA.UCase(VBA.Right(Data, 1))
                    If VBA.AscW(temp) > 64 And VBA.AscW(temp) < 91 Then
                        Keybd_Event_Combine VK_ALT, Asc(temp)
                        Flag = 1
                    End If
                End If
            End If
            
            
            If Flag = 0 Then
                State = GetKeyState(vbKeyCapital)
                If State = 1 Then '关闭Caps Lock
                    SetKeyState vbKeyCapital
                    Sleep 10
                End If
                
                For i = 1 To Len(Data)
                    MyData = VBA.Mid(Data, i, 1)
                    Select Case MyData
                        Case "!"
                            Keybd_Event_Combine VK_SHIFT, Asc("1")
                        Case "@"
                            Keybd_Event_Combine VK_SHIFT, Asc("2")
                        Case "#"
                            Keybd_Event_Combine VK_SHIFT, Asc("3")
                        Case "$"
                            Keybd_Event_Combine VK_SHIFT, Asc("4")
                        Case "%"
                            Keybd_Event_Combine VK_SHIFT, Asc("5")
                        Case "^"
                            Keybd_Event_Combine VK_SHIFT, Asc("6")
                        Case "&"
                            Keybd_Event_Combine VK_SHIFT, Asc("7")
                        Case "*"
                            Keybd_Event_Combine VK_SHIFT, Asc("8")
                        Case "("
                            Keybd_Event_Combine VK_SHIFT, Asc("9")
                        Case ")"
                            Keybd_Event_Combine VK_SHIFT, Asc("0")
                        Case "="
                            Keybd_Event_Data 187
                        Case "+"
                            Keybd_Event_Combine VK_SHIFT, 187 'Keybd_Event_Data VK_ADD
                        Case "-"
                            Keybd_Event_Data 189 'Keybd_Event_Data VK_SUBTRACT
                        Case "_"
                            Keybd_Event_Combine VK_SHIFT, 189
                        Case "."
                            Keybd_Event_Data VK_DECIMAL
                        Case "/"
                            Keybd_Event_Data VK_DIVIDE
                        Case "\"
                            Keybd_Event_Data 220
                        Case "|"
                            Keybd_Event_Combine VK_SHIFT, 220
                        Case ","
                            Keybd_Event_Data 188
                        Case "<"
                            Keybd_Event_Combine VK_SHIFT, 188
                        Case "."
                            Keybd_Event_Data 190
                        Case ">"
                            Keybd_Event_Combine VK_SHIFT, 190
                        Case "/"
                            Keybd_Event_Data 191
                        Case "?"
                            Keybd_Event_Combine VK_SHIFT, 191
                        Case "["
                            Keybd_Event_Data 219
                        Case "]"
                            Keybd_Event_Data 221
                        Case "{"
                            Keybd_Event_Combine VK_SHIFT, 219
                        Case "}"
                            Keybd_Event_Combine VK_SHIFT, 221
                        Case ";"
                            Keybd_Event_Data 186
                        Case ":"
                            Keybd_Event_Combine VK_SHIFT, 186
                        Case "'"
                            Keybd_Event_Data 222
                        Case """"
                            Keybd_Event_Combine VK_SHIFT, 222
                        Case "`"
                            Keybd_Event_Data 192
                        Case "~"
                            Keybd_Event_Combine VK_SHIFT, 192
                        Case Else
                            AscData = VBA.Asc(MyData)
                            If AscData >= 65 And AscData <= 90 Then
                                Keybd_Event_Combine VK_SHIFT, AscData
                            ElseIf AscData >= 97 And AscData <= 122 Then
                                AscData = AscData - 32
                                Keybd_Event_Data AscData
                            Else
                                Keybd_Event_Data AscData
                            End If
                    End Select
                    
                    Sleep 10
                        
                    DoEvents
                    
                Next i
            End If
    End Select
    
    Sleep 10
    
    DoEvents
    
    Call Cursor_Change
    
     
End Sub

'2014-12-24
'-------------------------------
Private Function Keybd_Event_Data(ByVal Data As Long)
    keybd_event Data, MapVirtualKey(Data, 0), 0, 0
    Sleep 10
    keybd_event Data, MapVirtualKey(Data, 0), KEYEVENTF_KEYUP, 0
End Function
Private Function Keybd_Event_Combine(ByVal Data1 As Long, ByVal Data2 As Long)
    keybd_event Data1, MapVirtualKey(Data1, 0), 0, 0
    keybd_event Data2, MapVirtualKey(Data2, 0), 0, 0
    keybd_event Data2, MapVirtualKey(Data2, 0), KEYEVENTF_KEYUP, 0
    keybd_event Data1, MapVirtualKey(Data1, 0), KEYEVENTF_KEYUP, 0
End Function

Private Function SetKeyState(ByVal Key As Long)
 keybd_event Key, MapVirtualKey(Key, 0), KEYEVENTF_EXTENDEDKEY Or 0, 0
 keybd_event Key, MapVirtualKey(Key, 0), KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Function

Private Sub test_Ctrl_Wheel()
    
    Init
    
    Ctrl_Wheel 0
    
End Sub

Public Function Ctrl_Wheel(Times As Integer, Optional Flag As Integer = 1) 'Ctrl+Wheel
Dim i As Integer

keybd_event VK_CTRL, MapVirtualKey(VK_CTRL, 0), 0, 0

If Flag = 1 Then '放大

    For i = 1 To Times
        mouse_event MOUSEEVENTF_WHEEL, 0, 0, 100, 0 '120为单击滚轮
        DoEvents
        Sleep 10
    Next
    
ElseIf Flag = -1 Then '缩小

    For i = 1 To Times
        mouse_event MOUSEEVENTF_WHEEL, 0, 0, -100, 0 '120为单击滚轮
        DoEvents
        Sleep 20
    Next
    
End If

keybd_event VK_CTRL, MapVirtualKey(VK_CTRL, 0), KEYEVENTF_KEYUP, 0

Sleep 20

End Function

Private Function MouseClick(x As Integer, y As Integer, Optional tWinsName As String)
Dim hwnd As Long
Dim ws As Object

    hwnd = FindWindow(vbNullString, tWinsName)
    If hwnd > 0 Then
        Set ws = CreateObject("WSCRIPT.SHELL")
        ws.AppActivate tWinsName
    End If
    
    mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
    Sleep 20
    mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
End Function
 

Public Function Cursor_Change()
Dim h1 As Long
Dim h2 As Long

h1 = GetCursor

Do While h1 = 65543 Or h1 = 65557 ' 66339 '65557
    DoEvents
    h1 = GetCursor
Loop


End Function

Public Function ws_PopUp(temp As String)
Dim ws As Object

Set ws = CreateObject("Wscript.SHELL")

ws.Popup temp, , "Prompt", 64

End Function


Public Function GetESC() As Boolean
    If GetKeyState(vbKeyEscape) Then
        GetESC = True
    Else
        GetESC = False
    End If
    
    DoEvents
    
End Function

Public Function SetESC()
    SetKeyState vbKeyEscape
    DoEvents
End Function

 
Private Function GetWindowName(hwnd As Long) As String
Dim lpBuffer As String * 256
Dim dwWindowCaption As String
Dim lpLength As Long

    lpLength = GetWindowText(hwnd, lpBuffer, 255)
    dwWindowCaption = VBA.Left(lpBuffer, lpLength)
    GetWindowName = dwWindowCaption

End Function


Public Function GetRange(myrow As Integer, x0 As Integer, y0 As Integer) As ScreenRange
    Dim MyRange As ScreenRange
    
    MyRange.Left = x0
    MyRange.Width = 77 'Default
    MyRange.Height = 17  'Default
    
    Select Case myrow
        Case 1
            MyRange.Top = y0 + 2
        Case 2
            MyRange.Top = y0 + 21
        Case 3
            MyRange.Top = y0 + 40
        Case 4
            MyRange.Top = y0 + 59
        Case 5
            MyRange.Top = y0 + 78
        Case 6
            MyRange.Top = y0 + 97
        Case 7
            MyRange.Top = y0 + 116
    End Select
 
    GetRange = MyRange
    
End Function

Public Function Init()
Dim temp As String
Dim Data() As String
Dim i As Integer
Dim hwnd As Long

    ReDim Data(5)
    Data(0) = "RIMSII - Order Management - Connected To: Real RIMS2 (AD Hong Kong)"
    Data(1) = "RIMSII - Order Management - Connected To: Real RIMS2 (AD Hong Kong) - [Purchase Orders View]"
    Data(2) = "RIMSII - Order Management - Connected To: Real RIMS2 (AD Nansha)"
    Data(3) = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong)"
    Data(4) = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong) - [Purchase Orders View]"
    Data(5) = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Nansha)"
    
    hwnd = 0
    For i = 0 To UBound(Data)
        hwnd = FindWindow(vbNullString, Data(i))
        If hwnd <> 0 Then
            temp = Data(i)
            Exit For
        End If
    Next i
    
        If hwnd = 0 Then
            ws_PopUp "没有找到 RIMSII-Order Management"
            End
        End If
    
    AppData = temp
    setHwnd = hwnd '句柄赋值到public变量
    
    Dim ws As Object
    Set ws = CreateObject("WSCRIPT.SHELL")
    ws.AppActivate temp
    
    If GetESC Then SetESC
    
End Function

Public Function GetCol(TitleStr, Optional Dict As Dictionary, Optional myrow) As String()
    Dim temp As String
    Dim Data() As String
    Dim Sdata() As String
    Dim i%, t%
    Dim r1 As Range
    
        Set Dict = CreateObject("scripting.dictionary")
        Dict.CompareMode = TextCompare
        
        temp = TitleStr
        Data = VBA.Split(temp, ",")
        t = 0
        For i = 0 To UBound(Data)
            temp = VBA.Trim(Data(i))
            Set r1 = Cells.Find(temp, , , xlPart)
            If Not r1 Is Nothing Then
                ReDim Preserve Sdata(t)
                Sdata(t) = r1.Column
                t = t + 1
                Dict(Data(i)) = r1.Column
            Else
                Debug.Print "Can not find :" & temp
            End If
        Next
        myrow = r1.Row
        GetCol = Data
    
End Function

