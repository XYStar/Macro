Attribute VB_Name = "B_Function_For_RIMSII_PO"
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type Size
        cx As Long
        cy As Long
End Type

Private Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Private Declare Function GetFocus Lib "user32" () As Long


Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long

'-----------------------------------------

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

'获取dwExtraInfo
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
'这个函数模拟了键盘行动  这个函数支持屏幕捕获（截图）。
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

'----------------------------------------


Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hwnd As Long, pt As POINTAPI, ByVal un As Long) As Long


'------------------------------------------------------------------

'枚举窗口列表中的所有父窗口 (顶级和被所有窗口)
'返回值 Long，非零表示成功，零表示失败
'lpEnumFunc Long，指向为每个子窗口都调用的一个函数的指针。用AddressOf运算符获得函数在标准模式下的地址
'lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
'为指定的父窗口枚举子窗口
'返回值 Long，非零表示成功，零表示失败
'hWndParent Long，欲枚举子窗口的父窗口的句柄
'lpEnumFunc Long，为每个子窗口调用的函数的指针。用AddressOf运算符获得函数在一个标准模块中的地址
'lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的。（原文：Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.）
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


Private Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'------------------------------------------------------------------
'寻找窗口列表中第一个符合指定条件的顶级窗口（在vb里使用：FindWindow最常见的一个用途是获得ThunderRTMain类的隐藏窗口的句柄；该类是所有运行中vb执行程序的一部分。获得句柄后，可用api函数GetWindowText取得这个窗口的名称；该名也是应用程序的标题）
'lpClassName  String，指向包含了窗口类名的空中止（C语言）字串的指针；或设为零，表示接收任何类
'lpWindowName  String，指向包含了窗口文本（或标签）的空中止（C语言）字串的指针；或设为零，表示接收任何窗口标题
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'在窗口列表中寻找与指定条件相符的第一个子窗口
'hWnd1  Long，在其中查找子的父窗口。如设为零，表示使用桌面窗口（通常说的顶级窗口都被认为是桌面的子窗口，所以也会对它们进行查找）
'hWnd2  Long，从这个窗口后开始查找。这样便可利用对FindWindowEx的多次调用找到符合条件的所有子窗口。如设为零，表示从第一个子窗口开始搜索
'lpsz1  String，欲搜索的类名。零表示忽略
'lpsz2  String，欲搜索的类名。零表示忽略
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'----------------------------------------
'判断一个窗口是否为另一窗口的子或隶属窗口
Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
'判断指定的句柄是否为一个菜单的句柄
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
'判断一个矩形是否为空
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long

'判断一个窗口句柄是否有效
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'判断窗口是否处于活动状态（在vb里使用：针对vb窗体和控件，请用enabled属性）
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

'----------------------------------------------------------------------------------------------
'获得一个窗口的句柄，该窗口与某源窗口有特定的关系
'返回值  Long，由wCmd决定的一个窗口的句柄。如没有找到相符窗口，或者遇到错误，则返回零值。会设置GetLastError
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'GetWindow()  wcmd
Private Const GW_HWNDFIRST = 0 '为一个源子窗口寻找第一个兄弟（同级）窗口，或寻找第一个顶级窗口
Private Const GW_HWNDLAST = 1 '为一个源子窗口寻找最后一个兄弟（同级）窗口，或寻找最后一个顶级窗口
Private Const GW_HWNDNEXT = 2 '为源窗口寻找下一个兄弟窗口
Private Const GW_HWNDPREV = 3 '为源窗口寻找前一个兄弟窗口
Private Const GW_OWNER = 4 '寻找窗口的所有者
Private Const GW_CHILD = 5 '寻找源窗口的第一个子窗口


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'GetWindowDC 获取整个窗口（包括边框、滚动条、标题栏、菜单等）的设备场景
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

'GetWindowExtEx 获取指定设备场景的窗口范围
Private Declare Function GetWindowExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As Size) As Long

'GetWindowLong 从指定窗口的结构中取得信息
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'GetWindowLong()
Private Const GWL_WNDPROC = (-4) '该窗口的窗口函数的地址
Private Const GWL_HINSTANCE = (-6) '拥有窗口的实例的句柄
Private Const GWL_HWNDPARENT = (-8) '该窗口之父的句柄。不要用SetWindowWord来改变这个值
Private Const GWL_STYLE = (-16) '窗口样式
Private Const GWL_EXSTYLE = (-20) '扩展窗口样式
Private Const GWL_USERDATA = (-21) '含义由应用程序规定
Private Const GWL_ID = (-12) '对话框中一个子窗口的标识符

Private Const DWL_MSGRESULT = 0 '在对话框函数中处理的一条消息返回的值
Private Const DWL_DLGPROC = 4 '这个窗口的对话框函数地址
Private Const DWL_USER = 8 '含义由应用程序规定


'GetWindowText 取得一个窗体的标题（caption）文字，或者一个控件的内容（在vb里使用：使用vb窗体或控件的caption或text属性）
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'GetWindowsDirectory 这个函数能获取Windows目录的完整路径名。在这个目录里，保存了大多数windows应用程序文件及初始化文件
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'GetWindowTextLength 调查窗口标题文字或控件内容的长短（在vb里使用：直接使用vb窗体或控件的caption或text属性）
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'GetWindowOrgEx 获取指定设备场景的逻辑窗口的起点
Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lppoint As POINTAPI) As Long

'GetWindowThreadProcessId 获取与指定窗口关联在一起的一个进程和线程标识符
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


'----------------------------------------------------------------------------------------------

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'==============================          GetMenu              =======================================
 
'GetMenu
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'hwnd Long，窗口的句柄
'bRevert Long，如设为TRUE，表示接收原始的系统菜单
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'hMenu Long，菜单的句柄
'nPos Long，条目在菜单中的位置。第一个条目的编号为0
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

'hMenu Long，菜单句柄
'wIDItem Long，欲接收的菜单条目的标识符。如果在wFlags参数中设置了MF_BYCOMMAND标志，这个参数就用于指定要改变的菜单条目的命令ID。如果设置的是MF_BYPOSITION标志，这个参数就用于指定条目在菜单中的位置（第一个条目的位置为0）
'lpString String，指定一个预先定义好的字串缓冲区，以便为菜单条目装载字串
'nMaxCount Long，载入lpString缓冲区中的最大字符数量+1
'wFlag Long，常数MF_BYCOMMAND或MF_BYPOSITION，取决于wID参数的设置

Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long


Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&

'hMenu Long，菜单的句柄
'un Long，菜单条目的菜单ID或位置
'b Boolean，如un指定的是条目位置，就为TRUE；如指定的是一个菜单ID，则为FALSE
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    Wid As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


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


'在一个矩形中装载指定菜单条目的屏幕坐标信息
'返回值 Long，TRUE（非零）表示成功，否则返回零。会设置GetLastError
'hWnd Long，包含指定菜单或弹出式菜单的一个窗口的句柄
'hMenu Long，菜单的句柄
'uItem Long，欲检查的菜单条目的位置或菜单ID
'lprcItem RECT，在这个结构中装载菜单条目的位置及大小（采用屏幕坐标表示）
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
'获取控件类型
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102

Private ChildHwnd As String   '缓存子窗体控件句柄


'===========2016============

Private Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Const MK_LBUTTON = &H1
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const BM_CLICK = &HF5


Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long


Dim XY As POINTAPI
Dim AppData As String

'===========2016============


'这是一个回调函数, 必须放在模块中. 用来遍历指定窗口的子窗口(控件). 这里参数中的 hWnd 即为子窗口(控件)句柄
Private Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
ChildHwnd = ChildHwnd & "," & hwnd
EnumChildProc = 1
End Function


' 函数: FGetClassName
' 功能: 返回指定窗口中的类型
' 参数: hWnd 指定窗口的句柄
' 返回: 指定窗口的类型

Private Function FGetClassName(hwnd As Long) As String
Dim ClassName As String
Dim Ret As Long
'填充缓冲(如果填充太小返回会不完整).
 ClassName = Space(256)

 '调用 GetClassName 函数, 返回值为类型名的实际长度.
 Ret = GetClassName(hwnd, ClassName, 256)
 
 '函数返回类型. Ret 为上一步所得到的类型名的实际长度.
 FGetClassName = Left(ClassName, Ret)
End Function

' 函数: GetText
' 功能: 返回指定窗口(如文本框)中的文字
' 参数: WindowHandle 指定窗口的句柄
' 返回: 指定窗口的文字
Private Function GetText(WindowHandle As Long) As String
Dim strBuffer As String '字符串缓冲
Dim Char As String '储存密码掩码以待恢复

    '填充缓冲(如果填充太小返回会不完整).
    strBuffer = VBA.Space(255)
    '发送消息 EM_GETPASSWORDCHAR(返回密码掩码) 给指定窗口. 这里返回掩码给Char(比如可能 Char=*).
    Char = SendMessage(WindowHandle, &HD2, 0, 0)
    '发送消息 EM_SETPASSWORDCHAR(设置密码掩码) 给指定窗口. 这里设置了0(Null), 即除去除密码掩码.
    PostMessage WindowHandle, &HCC, 0, 0
    '如果是Edit控件则等待消息发送成功, 即等待掩码被去除.
    If InStr("Edit", FGetClassName(WindowHandle)) And Char <> "0" Then Sleep (10)
    '发送消息 WM_GETTEXT(返回所含文字) 给指定窗口. 这里得到Edit控件的文字, 即密码. 注意"ByVal".
    SendMessage WindowHandle, &HD, 255, ByVal strBuffer
    '发送消息 EM_SETPASSWORDCHAR(设置密码掩码) 给指定窗口. 这里设置为Char, 即恢复原先掩码.
    PostMessage WindowHandle, &HCC, ByVal Char, 0
    '函数返回所得文字(密码), 之所以要用Trim去空格是因为第一步中用空格填充了255个字符.
    GetText = VBA.Replace(VBA.Trim(strBuffer), VBA.Chr(0), "")
End Function


''为指定的父窗口枚举子窗口
''返回值 Long，非零表示成功，零表示失败
''hWndParent Long，欲枚举子窗口的父窗口的句柄
''lpEnumFunc Long，为每个子窗口调用的函数的指针。用AddressOf运算符获得函数在一个标准模块中的地址
''lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的。（原文：Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.）
'Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal LPARAM As Long) As Long
'由于上面这个函数每次调用都会得到下一个子窗体（控件）的句柄，并赋值给hWnd,实际使用中，我把所有子句柄存放在ChildHwnd字符串中，遍历完毕，再
'Dim AllHwnd() As String
'去除多余的无效字符
'ChildHwnd =vba. Mid(ChildHwnd, 2)
'转换成数组
'AllHwnd =vba. Split(ChildHwnd, ",")

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

    Dim lpBuffer As String * 1024
    Dim dwWindowCaption As String
    Dim lpLength As Long

    lpLength = GetWindowText(hwnd, lpBuffer, 1024)
    dwWindowCaption = VBA.Left(lpBuffer, lpLength)
'    MsgBox dwWindowCaption
    Debug.Print dwWindowCaption

    If InStr(dwWindowCaption, "Word") > 0 Then
        '停止查找函数返回0
        EnumWindowsProc = 0
    Else
        '继续查找函数返回1
        EnumWindowsProc = 1
    End If

End Function
 
 

Private Function test_RIMSII_PO()

PO_OrderNO "3197566"
PO_Vendor "*"
Save_Form_PO
Cells(12, 6) = Get_PONO
PO_Comments 12
'PO_DueDate

End Function

Public Function Get_PONO() As String
 
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
        '去除多余的无效字符
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        '转换成数组
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------
    

Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim Path As String
Dim PicStr As String
Dim FormStr As String
Dim Flag As Integer

        Path = "C:\1\1.bmp"
        If Dir("C:\1\", vbDirectory) = "" Then
            MkDir "C:\1"
        End If
    '--------------
        Myhwnd = CLng(AllHwnd(26))
        SetFocusAPI Myhwnd
 

        'Ctrl+Wheel 放大
        '==============
            DoEvents
            Ctrl_Wheel 1
            DoEvents
            Sleep 200
        '===============
    
    Flag = 0
    Do
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '获取对应Rectangle坐标,Order No.的矩形区域
        '===================
            x0 = MyRect.Left + 18 'PONO.X
            y0 = MyRect.Top + 6 'PONO.Y
            GetPrintScreen Path, x0, y0, 79, 17
            PicStr = Image_Str(Path)
        '===================
        temp = PicStr
        temp = VBA.Replace(temp, VBA.Chr(0), "")
        If temp = "Image is not readable" Or VBA.Trim(temp) = "" Then
            PicStr = ""
            Flag = Flag + 1 '如果提取的内容为空,则多提取一次
        End If
        
        PicStr = VBA.Replace(PicStr, VBA.Chr(0), "")
        PicStr = VBA.Replace(PicStr, " ", "")
        PicStr = VBA.Replace(PicStr, VBA.Chr(10), "")
        PicStr = VBA.Replace(PicStr, VBA.Chr(13), "")
        Get_PONO = PicStr
        
        Debug.Print "PONO:" & PicStr
        
        If PicStr <> "" Then
            Exit Do
        End If
        
        If Flag = 2 Then '如果两次结果都为空，则跳出循环
            Exit Do
        End If
    Loop
    '--------------
    
    'Ctrl+Wheel 缩小,返回原来大小
    '==============
        DoEvents
        Ctrl_Wheel 1, -1
        DoEvents
        Sleep 100
    '===============
        
        Sleep 800
        DoEvents
        
End Function
 
Private Sub test_PO_OrderNO()
    
    PO_OrderNO "3205922", ""
    
End Sub

Public Function PO_OrderNO(OrderNO, Item)
    
    'Dim OrderNO As String
    'OrderNO = "3197566"
    
Dim hwnd As Long
Dim AllHwnd() As String

Dim i As Integer
Dim t As Long
Dim EditHwnd1 As Long
Dim temp As String
Dim StrData As String

    
        Call Init
        hwnd = setHwnd
        
        '-------------------
            ChildHwnd = ""
            Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
            ChildHwnd = VBA.Mid(ChildHwnd, 2)
            AllHwnd = VBA.Split(ChildHwnd, ",")
        '---------------------

Dim Myhwnd As Long
Dim ClassName As String

        For t = 0 To UBound(AllHwnd)
            
            Myhwnd = CLng(AllHwnd(t))
            ClassName = FGetClassName(Myhwnd)
            'Debug.Print ClassName
            'Debug.Print GetText(Myhwnd)
        
            Select Case LCase(ClassName)
                Case "button"

                    If GetText(Myhwnd) = "&Reset" Then
                        'Dim OrderNO
                        'OrderNO = "3197566"
                        temp = OrderNO
                        For i = 1 To Len(temp)
                            StrData = Mid(temp, i, 1)
                            PostMessage EditHwnd1, WM_CHAR, Asc(StrData), 0
                        Next i
        
                        hwnd = GetParent(EditHwnd1)
        
                        Sleep 100
                        PostMessage hwnd, WM_KEYDOWN, VK_ENTER, 0
                        PostMessage hwnd, WM_KEYUP, VK_ENTER, 0
                        
                        Delay 500
                        
                        Exit For
                    End If
                    
                Case "edit"
                    EditHwnd1 = Myhwnd
        
            End Select
        
        Next t
        
        '----------
        '选Item
            Call Sales_Order_Popup_Search_Window(Item)
        '---------

End Function

Public Function PO_Vendor(Vendor As String)

    Init
    SendMyData Vendor
    SendMyData "Enter"
    Delay 100
    
End Function

Public Function Save_Form_PO()

    Call Init
    SendMyData "^s"
    Delay 1000
    
End Function

Private Sub test_Sales_Order_Popup_Search_Window()
    Sales_Order_Popup_Search_Window "LEV-S68712"
End Sub

Private Function Sales_Order_Popup_Search_Window(Item)
 
Dim hwnd As Long
Dim AllHwnd() As String

Dim i As Integer
Dim j As Integer
Dim t As Long
Dim Data
Dim temp As String
Dim StrData As String
Dim Myhwnd As Long
Dim MyRect As RECT
Dim Flag As Integer

        
        '------------
        Dim Title As String
        If GetESC Then SetESC
        
        Title = "Sales Order Popup Search Window"
        hwnd = FindWindow(vbNullString, Title)
        t = 0
        Do While hwnd = 0
            DoEvents
            If GetESC Then
                Debug.Print "按ESC停止程序,是因为没有找到:" & Title
                End
            End If
            hwnd = FindWindow(vbNullString, Title)
            Delay 100
            t = t + 1
            If t > 15 Then '循环1.5秒
                Debug.Print "没有打开:" & Title
                Exit Function
            End If
        Loop
    
    
        Set ws = CreateObject("WSCRIPT.SHELL")
        ws.AppActivate Title
        Sleep 300
        '------------
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        '去除多余的无效字符
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        '转换成数组
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------
        'Debug.Print UBound(AllHwnd)

Dim x0, y0, x1 As Long, y1 As Long
Dim Path As String
Dim PicStr As String
Dim FormStr As String
Dim MyColorNum
Dim hDC As Long


        Path = "C:\1\1.bmp"
        If Dir("C:\1\", vbDirectory) = "" Then
            MkDir "C:\1"
        End If
    '--------------
        Myhwnd = CLng(AllHwnd(20))
        
        SetFocusAPI Myhwnd '获取焦点
        
        'Ctrl+Wheel 放大
        '==============
            DoEvents
            Ctrl_Wheel 4
            DoEvents
            Sleep 100
        '===============


        Call GetWindowRect(Myhwnd, MyRect) '获取对应Rectangle坐标,Order No.的矩形区域
        '===================
            x0 = MyRect.Left + 2
            y0 = MyRect.Top + 19
            x1 = MyRect.Right - 10
            y1 = MyRect.Bottom - 10
        '===================
 
        '----
        'Dim Item As String
        ReDim Data(0)
        'Item = "LEV-S41066"
        '###########################
        '选择Item ID 算法
            Do
                
                ws.AppActivate Title
                
                DoEvents
                
                '选中第一行
                '----
                SetCursorPos x0 + 10, y0 + 10
                MouseClick Title, x0 + 10, y0 + 10
                '----
                
                If Item = "" Then
                    Exit Do
                End If
                
                Flag = 0
                GetPrintScreen Path, x0, y0, 148, 108
                PicStr = Image_Str(Path, OneNote)
                
                If PicStr = "Image is not readable" Then
                    MsgBox "程序执行不成功,请重新尝试!", , "Image is not readable"
                    End
                End If
                
                Erase Data
                Data = VBA.Split(PicStr, VBA.Chr(10))
                For i = 0 To UBound(Data)
                    StrData = VBA.Replace(Data(i), " ", "")
                    Data(i) = StrData
                    Debug.Print StrData
                    If StrData <> "" Then
                        If Item = StrData Then
                            Flag = i + 1
                        End If
                    End If
                Next i
                '----
                
                If Flag > 0 Then
                    For i = 1 To Flag - 1
                        SendMyData "Down"
                    Next i
                    Sleep 300
                    Exit Do
                Else
                
                    '--------
                    hDC = GetDC(0)
                    MyColorNum = GetPixel(hDC, x1, y1) '判断右下角是否有滚动条出现
                    DeleteDC hDC
                    '--------
                    
                    If MyColorNum <> 16777215 Then '表示有滚动条出现，即需要按PageDown翻页
                        SendMyData "PgDn"
                        Sleep 200
                    Else
                        Debug.Print "没有找到Item:" & Item
                        MsgBox "没有找到Item:" & Item
                        End
                    End If
                End If
            Loop
        '###########################
        
        
        'Select
        '----
            SendMyData "%s"
            PO_Warning '判断是否有对话框弹出
            Sleep 500
        '----
        
End Function



Private Function My_Set() As Boolean
    '用于测试
    '============
    '2016-01-29
        Dim sys_infm, UserName, Pcname
        Set sys_infm = CreateObject("WSCRIPT.NETWORK")
        UserName = sys_infm.UserName
        Pcname = sys_infm.computername
        
        My_Set = False
                
        '此手提调试时，需延时处理
        If Pcname = "NBW7NANCS000020" Then
            My_Set = True
        End If
 
    '============
End Function

Private Function PO_Warning()

Dim hwnd As Long
Dim i As Integer

    hwnd = FindWindow(vbNullString, "Warning")
    i = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Exit Function
        End If
        
        If i = 10 Then
            Exit Do
        End If
        hwnd = FindWindow(vbNullString, "Warning")
        Delay 50
        i = i + 1
    Loop
    
    If hwnd <> 0 Then
        If My_Set Then
            SendMyData "%y"
        Else
            ws_PopUp "PO found for this Order Line!"
            End
        End If
    End If
    
End Function
 

Sub test_PO_Comments()
        
        Call Init
        
        PO_Comments 12
        
    'SendMyUnicode Cells(12, 4).Text
 
'    hwnd = FindWindow(vbNullString, "Comments ")
    'PO_Comments 12

'    Call ClearClipboard
'
'    Call SetClipboard(Cells(Trow, 4))
End Sub

Public Function PO_Comments(Trow As Integer)
Dim hwnd As Long
Dim i As Integer
Dim t As Integer
Dim MyRect As RECT
Dim Title As String
Dim temp As String
Dim StrData As String
Dim Data() As String

        Sleep 500

        
        '----------
        Call Init
        
        hwnd = setHwnd

    '-----------
        Sleep 100
        SendMyData "%o"
        SendMyData "c"
    '-----------
      
    Title = "Comments "
    hwnd = FindWindow(vbNullString, "Comments ")
    i = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Exit Function
        End If
        
        If i = 40 Then
            ws_PopUp "没有找到:Comments窗口!"
            End
        End If
        hwnd = FindWindow(vbNullString, "Comments ")
        Sleep 50
        i = i + 1
    Loop
    
    
    'Print
    '------------
    Call GetWindowRect(hwnd, MyRect)
    Call SetCursorPos(MyRect.Left + 350, MyRect.Top + 48)
    Call MouseClick(Title, MyRect.Left + 60, MyRect.Top + 20)
    Sleep 100
    '------------

        For i = 1 To 5
            SendMyData "Tab"
        Next i
        
        Dim r1  As Range
        Set r1 = Cells.Find("Comment", , , xlWhole)
        If r1 Is Nothing Then
            MsgBox "没有找到Comment列!"
            End
        End If
        temp = Cells(Trow, r1.Column).Text
        SendMyUnicode temp

        Sleep 100
        SendMyData "%s"
        SendMyData "%c"

End Function


Private Sub test_PO_DueDate()
   PO_DueDate "2016-3-4"
End Sub

Public Function PO_DueDate(DueDate As String)
Dim i As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String

    If DueDate = "" Then
        Exit Function
    End If

    Delay 500
    
    Call Init
    hwnd = setHwnd
    
    '------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------
    
Dim Myhwnd As Long

        Myhwnd = CLng(AllHwnd(24)) '测试可知，AllHwnd(24)为TabControl控件
        
        SetFocusAPI Myhwnd
        DoEvents
        For i = 1 To 3
            SendMyData "Right"
            Delay 100
        Next i
        
        Delay 500
        SendMyData "%f"
        Delay 500
        SendMyData DueDate
        Delay 100
        SendMyData "%o"
        Delay 100
        SendMyData "^s"
 
End Function

Public Function PO_New_Orders()
    Call Init
    
    SendMyData "^n"
    
    Delay 200
    
End Function

Private Function MouseClick(Title As String, x As Integer, y As Integer)
'Dim hwnd As Long
    hwnd = FindWindow(vbNullString, Title)
    If hwnd = 0 Then
        MsgBox "没有找到:" & Title
    Else
        Set ws = CreateObject("WSCRIPT.SHELL")
        ws.AppActivate Title
    End If
    
    mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
    Sleep 20
    mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
    
End Function
  
