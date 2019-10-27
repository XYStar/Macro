Attribute VB_Name = "Test_Use_For_Test"
Type POINTAPI
        x As Long
        y As Long
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type Size
        cx As Long
        cy As Long
End Type

Type WINDOWPLACEMENT
        Length As Long
        flags As Long
        showCmd As Long
        ptMinPosition As POINTAPI
        ptMaxPosition As POINTAPI
        rcNormalPosition As RECT
End Type

Declare Function GetFocus Lib "user32" () As Long

'-----------------------------------------

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

'获取dwExtraInfo
Declare Function GetMessageExtraInfo Lib "user32" () As Long
'这个函数模拟了键盘行动  这个函数支持屏幕捕获（截图）。
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move

Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

'----------------------------------------


Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Declare Function ChildWindowFromPoint Lib "user32" (ByVal hWndParent As Long, ByVal pt As POINTAPI) As Long
Declare Function ChildWindowFromPointEx Lib "user32" (ByVal hwnd As Long, pt As POINTAPI, ByVal un As Long) As Long


'------------------------------------------------------------------

'枚举窗口列表中的所有父窗口 (顶级和被所有窗口)
'返回值 Long，非零表示成功，零表示失败
'lpEnumFunc Long，指向为每个子窗口都调用的一个函数的指针。用AddressOf运算符获得函数在标准模式下的地址
'lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
'为指定的父窗口枚举子窗口
'返回值 Long，非零表示成功，零表示失败
'hWndParent Long，欲枚举子窗口的父窗口的句柄
'lpEnumFunc Long，为每个子窗口调用的函数的指针。用AddressOf运算符获得函数在一个标准模块中的地址
'lParam Long，在枚举期间，传递给dwcbkd32.ocx定制控件之EnumWindows事件的值。这个值的含义是由程序员规定的。（原文：Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.）
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'------------------------------------------------------------------
'寻找窗口列表中第一个符合指定条件的顶级窗口（在vb里使用：FindWindow最常见的一个用途是获得ThunderRTMain类的隐藏窗口的句柄；该类是所有运行中vb执行程序的一部分。获得句柄后，可用api函数GetWindowText取得这个窗口的名称；该名也是应用程序的标题）
'lpClassName  String，指向包含了窗口类名的空中止（C语言）字串的指针；或设为零，表示接收任何类
'lpWindowName  String，指向包含了窗口文本（或标签）的空中止（C语言）字串的指针；或设为零，表示接收任何窗口标题
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'在窗口列表中寻找与指定条件相符的第一个子窗口
'hWnd1  Long，在其中查找子的父窗口。如设为零，表示使用桌面窗口（通常说的顶级窗口都被认为是桌面的子窗口，所以也会对它们进行查找）
'hWnd2  Long，从这个窗口后开始查找。这样便可利用对FindWindowEx的多次调用找到符合条件的所有子窗口。如设为零，表示从第一个子窗口开始搜索
'lpsz1  String，欲搜索的类名。零表示忽略
'lpsz2  String，欲搜索的类名。零表示忽略
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_GETTEXT = &HD

'----------------------------------------
'判断一个窗口是否为另一窗口的子或隶属窗口
Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
'判断指定的句柄是否为一个菜单的句柄
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
'判断一个矩形是否为空
Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long

'判断一个窗口句柄是否有效
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'判断窗口是否处于活动状态（在vb里使用：针对vb窗体和控件，请用enabled属性）
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

'----------------------------------------------------------------------------------------------
'获得一个窗口的句柄，该窗口与某源窗口有特定的关系
'返回值  Long，由wCmd决定的一个窗口的句柄。如没有找到相符窗口，或者遇到错误，则返回零值。会设置GetLastError
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'GetWindow()  wcmd
Private Const GW_HWNDFIRST = 0 '为一个源子窗口寻找第一个兄弟（同级）窗口，或寻找第一个顶级窗口
Private Const GW_HWNDLAST = 1 '为一个源子窗口寻找最后一个兄弟（同级）窗口，或寻找最后一个顶级窗口
Private Const GW_HWNDNEXT = 2 '为源窗口寻找下一个兄弟窗口
Private Const GW_HWNDPREV = 3 '为源窗口寻找前一个兄弟窗口
Private Const GW_OWNER = 4 '寻找窗口的所有者
Private Const GW_CHILD = 5 '寻找源窗口的第一个子窗口


Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'GetWindowDC 获取整个窗口（包括边框、滚动条、标题栏、菜单等）的设备场景
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

'GetWindowExtEx 获取指定设备场景的窗口范围
Declare Function GetWindowExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As Size) As Long

'GetWindowLong 从指定窗口的结构中取得信息
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

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
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'GetWindowsDirectory 这个函数能获取Windows目录的完整路径名。在这个目录里，保存了大多数windows应用程序文件及初始化文件
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'GetWindowTextLength 调查窗口标题文字或控件内容的长短（在vb里使用：直接使用vb窗体或控件的caption或text属性）
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'GetWindowOrgEx 获取指定设备场景的逻辑窗口的起点
Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lppoint As POINTAPI) As Long

'GetWindowThreadProcessId 获取与指定窗口关联在一起的一个进程和线程标识符
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


'----------------------------------------------------------------------------------------------

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'==============================          GetMenu              =======================================

'GetMenu
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'hwnd Long，窗口的句柄
'bRevert Long，如设为TRUE，表示接收原始的系统菜单
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'hMenu Long，菜单的句柄
'nPos Long，条目在菜单中的位置。第一个条目的编号为0
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

'hMenu Long，菜单句柄
'wIDItem Long，欲接收的菜单条目的标识符。如果在wFlags参数中设置了MF_BYCOMMAND标志，这个参数就用于指定要改变的菜单条目的命令ID。如果设置的是MF_BYPOSITION标志，这个参数就用于指定条目在菜单中的位置（第一个条目的位置为0）
'lpString String，指定一个预先定义好的字串缓冲区，以便为菜单条目装载字串
'nMaxCount Long，载入lpString缓冲区中的最大字符数量+1
'wFlag Long，常数MF_BYCOMMAND或MF_BYPOSITION，取决于wID参数的设置

Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long


Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

'hMenu Long，菜单的句柄
'un Long，菜单条目的菜单ID或位置
'b Boolean，如un指定的是条目位置，就为TRUE；如指定的是一个菜单ID，则为FALSE
Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Type MENUITEMINFO
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

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


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
Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
'获取控件类型
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public ChildHwnd As String '缓存子窗体控件句柄


'===========2016============

Declare Function SetFocusAPI& Lib "user32" Alias "SetFocus" (ByVal hwnd As Long)
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Const MK_LBUTTON = &H1
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Private Const BM_CLICK = &HF5


Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lppoint As POINTAPI) As Long


Dim XY As POINTAPI
'===========2016============


'这是一个回调函数, 必须放在模块中. 用来遍历指定窗口的子窗口(控件). 这里参数中的 hWnd 即为子窗口(控件)句柄
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
ChildHwnd = ChildHwnd & "," & hwnd
EnumChildProc = 1
End Function


' 函数: FGetClassName
' 功能: 返回指定窗口中的类型
' 参数: hWnd 指定窗口的句柄
' 返回: 指定窗口的类型

Public Function FGetClassName(hwnd As Long) As String
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
Public Function GetText(WindowHandle As Long) As String
Dim strBuffer As String '字符串缓冲
Dim Char As String '储存密码掩码以待恢复

    '填充缓冲(如果填充太小返回会不完整).
    strBuffer = Space(255)

    '发送消息 EM_GETPASSWORDCHAR(返回密码掩码) 给指定窗口. 这里返回掩码给Char(比如可能 Char=*).
    Char = SendMessage(WindowHandle, &HD2, 0, 0)

    '发送消息 EM_SETPASSWORDCHAR(设置密码掩码) 给指定窗口. 这里设置了0(Null), 即除去除密码掩码.
    PostMessage WindowHandle, &HCC, 0, 0

    '如果是Edit控件则等待消息发送成功, 即等待掩码被去除.
    If InStr("Edit", FGetClassName(WindowHandle)) And Char <> "0" Then Sleep (10)

    '发送消息 WM_GETTEXT(返回所含文字) 给指定窗口. 这里得到Edit控件的文字, 即密码. 注意"ByVal", 如果少这个则VB崩溃.
    SendMessage WindowHandle, &HD, 255, ByVal strBuffer

    '发送消息 EM_SETPASSWORDCHAR(设置密码掩码) 给指定窗口. 这里设置为Char, 即恢复原先掩码.
    PostMessage WindowHandle, &HCC, ByVal Char, 0

    '函数返回所得文字(密码), 之所以要用Trim去空格是因为第一步中用空格填充了255个字符.
    GetText = Trim(strBuffer)
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

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim lpBuffer As String * 1024
    Dim dwWindowCaption As String
    Dim lpLength As Long

    lpLength = GetWindowText(hwnd, lpBuffer, 1024)
    dwWindowCaption = Left(lpBuffer, lpLength)
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

Sub myCursorPos()


'
    GetCursorPos XY
    Debug.Print XY.x
    Debug.Print XY.y
End Sub


Sub Test_RIMSII()

Dim lpEnumFunc As Long
Dim lParam As Long
Dim Data As String
Dim hwnd As Long
Dim ChWnd As Long
Dim t As Long
Dim AllHwnd() As String

Dim lpBuffer As String * 256
Dim dwWindowCaption As String
Dim lpLength As Long
Dim temp As String
Dim MyRect As RECT
Dim Title As String
Dim lnglen
Dim thwnd



hwnd = FindWindow(vbNullString, "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong)")
'Debug.Print GetFocus
'Sales Order Popup Search Window
hwnd = FindWindow(vbNullString, "Popup Search")


temp = "Popup Search"
'temp = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong)"
Set ws = CreateObject("WSCRIPT.SHELL")
ws.AppActivate temp

Title = temp
'----------------------------------------

Call GetWindowRect(hwnd, MyRect)

'Debug.Print myRect.Bottom
'Debug.Print myRect.Left
'Debug.Print myRect.Right
'Debug.Print myRect.Top


Dim x0 As Integer
Dim y0 As Integer
 

x0 = MyRect.Left
y0 = MyRect.Top

'Sleep 100
'
'PostMessage hwnd, WM_KEYDOWN, VK_TAB, 0
'PostMessage hwnd, WM_KEYUP, VK_TAB, 0
'


'Call GetClientRect(hwnd, myRect)
'
''Debug.Print myRect.Bottom
''Debug.Print myRect.Left
''Debug.Print myRect.Right
''Debug.Print myRect.Top
'
Dim x1 As Integer
Dim y1 As Integer
'
x1 = MyRect.Right
y1 = MyRect.Bottom



'x0 = x0 + 60
'y0 = y0 + 326 'Y+316-->Y+336 ,取中间值:326

'Call SetCursorPos(x0, y0)
'MouseClick temp, x0, y0
'Call SetCursorPos(x1, y1)
'
ChildHwnd = ""

'调用 EnumChidWindows 函数开始遍历指定窗口的子窗口(控件). 第一个参数即指定窗口的句柄, 第二个参数为所需回调函数的地址(由AddressOf操作符获得), 第三个参数不用管...
t = EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)

'去除多余的无效字符
ChildHwnd = VBA.Mid(ChildHwnd, 2)
'转换成数组
AllHwnd = VBA.Split(ChildHwnd, ",")

Data = ""

Dim Myhwnd As Long


'Button 按钮控件
'COMBOBOX 组合框控件
'Edit 编辑框控件
'ListBox 列表框控件
'ScrollBar 滚动条控件
'Static 静态控件

Dim ClassName As String


For t = 0 To UBound(AllHwnd)

    Myhwnd = CLng(AllHwnd(t))

        'SetFocusAPI Myhwnd

        'Sleep 500

    lpBuffer = ""
    lpLength = GetWindowText(Myhwnd, lpBuffer, 255)
    temp = FGetClassName(Myhwnd)

    ClassName = temp

    'Debug.Print temp
    'Debug.Print temp & " " & Myhwnd

    If lpLength <> 0 Then
        dwWindowCaption = Left(lpBuffer, lpLength)
        'Debug.Print dwWindowCaption & "  " & Myhwnd
        'Debug.Print Myhwnd & " " & dwWindowCaption & " " & temp & " " & t
    Else
        dwWindowCaption = ""
        'Debug.Print Myhwnd & " " & temp & " " & t
    End If
    
    If t = 3 Then
        Call GetWindowRect(Myhwnd, MyRect)
        x0 = MyRect.Left + 30 '33
        y0 = MyRect.Top + 16 '20
        x1 = MyRect.Right
        y1 = MyRect.Bottom
    
    
    Dim Path As String
        Do
            DoEvents
            Path = "F:\AVERY DENNISON\MyCode\New Code_A_Star\Z For RIMSII\1.bmp"
            GetPrintScreen x0, y0, Round((x1 - x0) / 4) - 45, (y1 - y0 - 10), Path '18  10
            
            Path = "F:\AVERY DENNISON\MyCode\New Code_A_Star\Z For RIMSII\1.bmp"
            temp = GetPictureText(Path)
            Debug.Print temp
            If temp <> "Image is not readable" Then
                Exit Do
            End If
        Loop
    
        Exit For
        
    End If
    
    'lnglen = PostMessage(Myhwnd, WM_GETTEXTLENGTH, 0, 0) 'WM_GETTEXT
    'Debug.Print lnglen
        
'    If t = 2 Then 'pbdw80
'        'Myhwnd
'
'        Call GetClientRect(Myhwnd, myRect)
'        Debug.Print myRect.Bottom
'        Debug.Print myRect.Left
'        Debug.Print myRect.Right
'        Debug.Print myRect.Top
'        x0 = myRect.Left
'        y0 = myRect.Top
'        x1 = myRect.Right
'        y1 = myRect.Bottom
'
'        Call SetCursorPos(x0, y0)
'        Call SetCursorPos(x1, y1)
'
'        lnglen = PostMessage(Myhwnd, WM_GETTEXT, 0, 0)
'        'Debug.Print lnglen
'
'    End If
'
'    If t = 4 Then '8258746 Edit
'        Call GetClientRect(Myhwnd, myRect)
'        Debug.Print myRect.Bottom
'        Debug.Print myRect.Left
'        Debug.Print myRect.Right
'        Debug.Print myRect.Top
'        x0 = myRect.Left
'        y0 = myRect.Top
'        x1 = myRect.Right
'        y1 = myRect.Bottom
'
'        Call SetCursorPos(x0, y0)
'        'Myhwnd
'        'lnglen = PostMessage(Myhwnd, WM_GETTEXT, 0, 0)
'        'Debug.Print lnglen
'    End If
'
'    If t = 5 Then '1709308 Edit
'        Call GetClientRect(Myhwnd, myRect)
'        Debug.Print myRect.Bottom
'        Debug.Print myRect.Left
'        Debug.Print myRect.Right
'        Debug.Print myRect.Top
'        x0 = myRect.Left
'        y0 = myRect.Top
'        x1 = myRect.Right
'        y1 = myRect.Bottom
'
'        Call SetCursorPos(x0, y0)
'        'Myhwnd
'        'lnglen = PostMessage(Myhwnd, WM_GETTEXT, 0, 0)
'        'Debug.Print lnglen
'    End If
    
        
'        Dim i As Integer
'        Dim Strdata As String
'        Dim b As Object
'
'            Select Case LCase(temp)
'                Case "button"
'                    'Debug.Print "button Caption: " & dwWindowCaption & " " & Myhwnd
'
'                    If dwWindowCaption = "&Close" Then
'                        '方式1:
'                        'PostMessage Myhwnd, BM_CLICK, 0, 0
'
'                        '方式2:
'                        'PostMessage Myhwnd, WM_LBUTTONDOWN, MK_LBUTTON, 0
'                        'PostMessage Myhwnd, WM_LBUTTONUP, MK_LBUTTON, 0
'
'        '----------------------------------
'                        '方式3:
'        '                SetFocusAPI Myhwnd
'        '                Sleep 20
'        '                hwnd = GetParent(Myhwnd)
'        '                PostMessage hwnd, WM_KEYDOWN, VK_ENTER, 0
'        '                PostMessage hwnd, WM_KEYUP, VK_ENTER, 0
'        '                Sleep 20
'        '----------------------------------
'                    End If
'
'            End Select
'
'            If t = 20 Then 'Item ID,测试可知当t=20，对应Item ID
'                Call GetWindowRect(Myhwnd, myRect) '获取Item ID对应Rectangle坐标
'                Call SetCursorPos(myRect.Left + 60, myRect.Top + 20)
'                Call MouseClick(Title, myRect.Left + 60, myRect.Top + 20)
'            End If

Next t



End Sub


Sub test_RIMSII_0222()

Dim lpEnumFunc As Long
Dim lParam As Long
Dim Data As String
Dim hwnd As Long
Dim ChWnd As Long
Dim t As Long
Dim AllHwnd() As String

Dim lpBuffer As String * 256
Dim dwWindowCaption As String
Dim lpLength As Long
Dim temp As String


hwnd = FindWindow(vbNullString, "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong)")
'Debug.Print GetFocus
'hwnd = FindWindow(vbNullString, "Sales Order Popup Search Window")

temp = "Sales Order Popup Search Window"
temp = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong)"
Set ws = CreateObject("WSCRIPT.SHELL")
ws.AppActivate temp

'----------------------------------------

'Sleep 100
'
'PostMessage hwnd, WM_KEYDOWN, VK_TAB, 0
'PostMessage hwnd, WM_KEYUP, VK_TAB, 0
'
'Exit Sub

ChildHwnd = ""

'调用 EnumChidWindows 函数开始遍历指定窗口的子窗口(控件). 第一个参数即指定窗口的句柄, 第二个参数为所需回调函数的地址(由AddressOf操作符获得), 第三个参数不用管...
t = EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)

'去除多余的无效字符
ChildHwnd = VBA.Mid(ChildHwnd, 2)
'转换成数组
AllHwnd = VBA.Split(ChildHwnd, ",")

Data = ""

Dim Myhwnd As Long


'Button 按钮控件
'COMBOBOX 组合框控件
'Edit 编辑框控件
'ListBox 列表框控件
'ScrollBar 滚动条控件
'Static 静态控件

Dim ClassName As String
Dim i As Integer
Dim StrData As String
Dim EditHwnd As Long

For t = 0 To UBound(AllHwnd)

    Myhwnd = CLng(AllHwnd(t))
    lpBuffer = ""
    lpLength = GetWindowText(Myhwnd, lpBuffer, 255)
    temp = FGetClassName(Myhwnd)
    ClassName = temp

    If lpLength <> 0 Then
        dwWindowCaption = Left(lpBuffer, lpLength)
    End If
    
    Debug.Print temp
    

    Select Case LCase(temp)
        Case "button"
            If dwWindowCaption = "&Reset" Then
                temp = "3121068"
                For i = 1 To Len(temp)
                    StrData = Mid(temp, i, 1)
                    PostMessage EditHwnd, WM_CHAR, Asc(StrData), 0   ' 发送一个 A 字符
                Next i

                hwnd = GetParent(EditHwnd)

                Sleep 100
                PostMessage hwnd, WM_KEYDOWN, VK_ENTER, 0
                PostMessage hwnd, WM_KEYUP, VK_ENTER, 0
            End If

        Case "edit"
            EditHwnd = Myhwnd

    End Select

'
'    Select Case LCase(temp)
'        Case "button"
'            Debug.Print "button Caption: " & dwWindowCaption & " " & Myhwnd
'
''            If dwWindowCaption = "&Close" Then
''
''                SetFocusAPI Myhwnd
''
''                Sleep 20
''                PostMessage hwnd, WM_KEYDOWN, VK_ENTER, 0
''                PostMessage hwnd, WM_KEYUP, VK_ENTER, 0
''                Sleep 20
''            End If
'
'        Case "edit"
'
'            Debug.Print "edit Caption: " & dwWindowCaption
'
'            'PostMessage Myhwnd, WM_CHAR, Asc("A"), 0
'
''            If Myhwnd = 328962 Then
''                temp = "3121068"
''                For i = 1 To Len(temp)
''                    Strdata = Mid(temp, i, 1)
''                    PostMessage Myhwnd, WM_CHAR, Asc(Strdata), 0   ' 发送一个 A 字符
''                Next i
''
''
''                hwnd = GetParent(Myhwnd)
''
''                Sleep 100
''                PostMessage hwnd, WM_KEYDOWN, VK_TAB, 0
''                PostMessage hwnd, WM_KEYUP, VK_TAB, 0
''            End If
'
''        Case "combobox"
''            Debug.Print "combobox"
'
''        Case "listbox"
''            Debug.Print "listbox"
''
''        Case "scrollbar"
''            Debug.Print "scrollbar"
''
''        Case "static"
''            Debug.Print "static"
'
'    End Select

Next t



End Sub
Private Function MouseClick(Title As String, x As Integer, y As Integer)
'Dim hwnd As Long
    hwnd = FindWindow(vbNullString, Title)
    If hwnd = 0 Then
        MsgBox "没有找到:" & Title
        End
    Else
        Set ws = CreateObject("WSCRIPT.SHELL")
        ws.AppActivate Title
    End If

    mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
    Sleep 20
    mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0

End Function


'TCHAR buf[1024];
'SendMessage(hwndEdit, WM_GETTEXT, sizeof(buf)/sizeof(TCHAR), (LPARAM)(void*)buf);

