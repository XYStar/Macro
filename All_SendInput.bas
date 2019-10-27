Attribute VB_Name = "All_SendInput"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As DEVMODE) As Long

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SRCCOPY = &HCC0020 ' (DWORD) destination = source
 
 
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

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2



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
 
 
 
'############################################################################
'############################################################################

'模拟按键，输入Unicode码---Backup 2015-06-17
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Type INPUT_TYPE
      dwType   As Long
      xi(0 To 23) As Byte
End Type
 
Private Type HARDWAREINPUT
    uMsg   As Long
    wParamL   As Integer
    wParamH   As Integer
End Type
  
'KEYBDINPUT
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const KEYEVENTF_UNICODE = &H4
Private Const KEYEVENTF_KEYDOWN = &H0
 
'INPUT_TYPE
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2

'###################################################
'2015-08-25
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Type GENERALINPUT
      dwType As Long
      xi(0 To 23) As Byte
End Type
Private Type KEYBDINPUT
     wVk As Integer
     wScan As Integer
     dwFlags As Long
     time As Long
     dwExtraInfo As Long
End Type
 
Dim AppData As String
 
Public Sub Test_MySendKey()
    Call Init

    'VK_SHIFT
    SendConbinedKey VK_SHIFT, VK_PRIOR
    
End Sub
'2015-08-25
Public Function SendCombinedKey(key1 As Long, key2 As Long)
Dim GInput(3) As GENERALINPUT
Dim KInput As KEYBDINPUT

    With KInput
        .wVk = key1
        .dwFlags = KEYEVENTF_KEYDOWN Or KEYEVENTF_UNICODE '按下键标志
    End With
    GInput(0).dwType = INPUT_KEYBOARD '表示该消息为键盘消息
    CopyMemory GInput(0).xi(0), KInput, Len(KInput) '把内存中KInput的数据复制到GInput
    
    
    
    With KInput
        .wVk = key2
        .dwFlags = KEYEVENTF_KEYDOWN Or KEYEVENTF_UNICODE '按下键标志
    End With
    GInput(1).dwType = INPUT_KEYBOARD '表示该消息为键盘消息
    CopyMemory GInput(1).xi(0), KInput, Len(KInput) '把内存中KInput的数据复制到GInput
    
    
    
    
    With KInput
        .wVk = key1  '要模拟的按键
        .dwFlags = KEYEVENTF_KEYUP Or KEYEVENTF_UNICODE '释放按键
    End With
    GInput(2).dwType = INPUT_KEYBOARD '表示该消息为键盘消息
    CopyMemory GInput(2).xi(0), KInput, Len(KInput) '把内存中KInput的数据复制到GInput
    
    
    
    With KInput
        .wVk = key2  '要模拟的按键
        .dwFlags = KEYEVENTF_KEYUP Or KEYEVENTF_UNICODE '释放按键
    End With
    GInput(3).dwType = INPUT_KEYBOARD '表示该消息为键盘消息
    CopyMemory GInput(3).xi(0), KInput, Len(KInput) '把内存中KInput的数据复制到GInput
    
    SendInput 4, GInput(0), Len(GInput(0))       '把GInput中存放的消息插入到消息列队

End Function
 
'2015-08-25
Public Function MySendKey(key1 As Long)
'参数bkey传入要模拟按键的虚拟码即可模拟按下指定键
Dim GInput(1) As GENERALINPUT
Dim KInput As KEYBDINPUT
 
'把按下键和释放键共2条键盘消息加入到GInput数据结构中
'===========================
    With KInput
        .wVk = key1  '要模拟的按键
        .dwFlags = KEYEVENTF_KEYDOWN Or KEYEVENTF_UNICODE '按下键标志
    End With
    
    GInput(0).dwType = INPUT_KEYBOARD '表示该消息为键盘消息
    CopyMemory GInput(0).xi(0), KInput, Len(KInput) '把内存中KInput的数据复制到GInput
'===========================
    With KInput
        .wVk = key1  '要模拟的按键
        .dwFlags = KEYEVENTF_KEYUP Or KEYEVENTF_UNICODE '释放按键
    End With
    
    GInput(1).dwType = INPUT_KEYBOARD '表示该消息为键盘消息
    CopyMemory GInput(1).xi(0), KInput, Len(KInput) '把内存中KInput的数据复制到GInput
'===========================

SendInput 2, GInput(0), Len(GInput(0))       '把GInput中存放的消息插入到消息列队

End Function



'发送中文
'####################################################################
'####################################################################
'2015-08-25
Public Function SendMyUnicode(ByVal theStr As String)
    Dim i As Long
    For i = 1 To Len(theStr)
        Call ProcessChar(VBA.Mid(theStr, i, 1))
        Sleep 10
        DoEvents
    Next
    
End Function

Private Function ProcessChar(this As String)
   Dim code As Long
   code = VBA.AscW(this)
   Call StuffBufferW(code)
End Function

Private Function StuffBufferW(ByVal CharCode As Long)
  Dim Retval As Long
  Dim IT As GENERALINPUT
  Dim KI(1) As KEYBDINPUT
  Dim i As Integer
  
        With KI(0)
            .wVk = 0
            .wScan = CharCode
            .dwFlags = KEYEVENTF_KEYDOWN Or KEYEVENTF_UNICODE
            .time = 0
            .dwExtraInfo = 0
        End With
        
        With KI(1)
            .wVk = 0
            .wScan = CharCode
            .dwFlags = KEYEVENTF_KEYUP Or KEYEVENTF_UNICODE
            .time = 0
            .dwExtraInfo = 0
        End With
        
        With IT
            .dwType = INPUT_KEYBOARD
        End With
            
        For i = 0 To UBound(KI)
            MoveMemory IT.xi(0), KI(i), Len(KI(i))
            Retval = SendInput(1&, IT, Len(IT))
            If Retval = 0 Then
                ws_PopUp "文字发送失败，请联系Star He!"
                End
            End If
        Next i

End Function
 





