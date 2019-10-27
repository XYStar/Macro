Attribute VB_Name = "Test_Excel_To_RIMSII_PO"
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
Dim setHwnd As Long
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
    strBuffer = Space(255)
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

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

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

Private Sub RIMSII_Order()

    If GetESC Then SetESC
    
    
Dim i As Integer
Dim j As Integer
Dim temp As String

Dim Endrow As Integer
Dim Erow As Integer
Dim Srow As Integer

    Endrow = Cells(60000, 1).End(xlUp).Row
    If Endrow < 12 Then
        ws_PopUp "没有数据!"
        Exit Sub
    End If
    
Dim OrderNO As String
Dim Vendor As String
Dim DueDate As String
Dim Comment As String

    Srow = Cells(8, 2)
    Erow = Cells(9, 2)
    
    If Srow = 0 Then
        Cells(8, 2) = 12
        Srow = 12
    End If
    
    If Erow = 0 Then
        Cells(9, 2) = Endrow
        Erow = Endrow
    End If
    
    For i = Srow To Erow
        '----------
            OrderNO = Cells(i, 1)
            Vendor = Cells(i, 2)
            Call RIMSII_OrderNO_Vender(OrderNO, Vendor)
            Cells(i, 1).Interior.Color = 3407718
            Cells(i, 2).Interior.Color = 3407718
        '----------
        
        '----------
            Call RIMSII_Comments(i)
            Cells(i, 4).Interior.Color = 3407718
        '----------
        
        '----------
            DueDate = Cells(i, 3)
            Call RIMSII_DueDate(DueDate)
            Cells(i, 3).Interior.Color = 3407718
        '----------
        
            Call RIMSII_New_Orders
        
        
        Cells(i, 5) = "Done:" & Now
    Next i
    
End Sub

Private Function Init()
Dim temp As String
Dim Data() As String
Dim i As Integer
 

    ReDim Data(2)
    Data(0) = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong)"
    Data(1) = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Hong Kong) - [Purchase Orders View]"
    Data(2) = "RIMSII - Purchase Order Management - Connected To: Real RIMS2 (AD Nansha)"
    
    hwnd = 0
    For i = 0 To UBound(Data)
        hwnd = FindWindow(vbNullString, Data(i))
        If hwnd <> 0 Then
            temp = Data(i)
            Exit For
        End If
    Next i
    
        If hwnd = 0 Then
            ws_PopUp "没有找到 RIMSII-Purchase Order Management"
            End
        End If
    
    AppData = temp
    setHwnd = hwnd
    
    Set ws = CreateObject("WSCRIPT.SHELL")
    ws.AppActivate temp
    
End Function

Private Function RIMSII_OrderNO_Vender(OrderNO As String, Vendor As String)

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
Dim lpBuffer As String * 256
Dim dwWindowCaption As String
Dim lpLength As Long
        
        For t = 0 To UBound(AllHwnd)
            
            Myhwnd = CLng(AllHwnd(t))
            temp = FGetClassName(Myhwnd)
            
            '--------
                lpBuffer = ""
                lpLength = GetWindowText(Myhwnd, lpBuffer, 255)
                If lpLength <> 0 Then
                    dwWindowCaption = Left(lpBuffer, lpLength)
                Else
                    dwWindowCaption = ""
                End If
            '--------
        
            Select Case LCase(temp)
                Case "button"
                    If dwWindowCaption = "&Reset" Then
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
            Call RIMSII_Sales_Order_Show
        
        '---------
        '输入Vendor
            SendMyData Vendor
            Delay 100
            
            SendMyData "^s"
            
            Delay 500
 
End Function
Private Function Get_ItemID() As String '2016-09-14 New
 
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
        
        SetFocusAPI Myhwnd
        
        'Ctrl+Wheel 放大
        '==============
            DoEvents
            Ctrl_Wheel 3
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
        Dim Item As String
        ReDim Data(0)
        Item = "LEV-S41066"
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
                
                Flag = 0
                Set_Picture Path, x0, y0, 150, 108
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
                    'Debug.Print StrData
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
            
            PO_Warning
            
            Sleep 800
        '----
        
        If Application.Version = 14 Then DeleteOneNotePages OneNote
        
        
End Function

Private Function Sales_Order_Popup_Search_Window()

Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim MyRect As RECT
Dim Title As String


    hwnd = FindWindow(vbNullString, "Sales Order Popup Search Window")
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Debug.Print "没有找到:Sales Order Popup Search Window"
            End
        End If
        hwnd = FindWindow(vbNullString, "Sales Order Popup Search Window")
    Loop

    Title = "Sales Order Popup Search Window"
    Set ws = CreateObject("WSCRIPT.SHELL")
    ws.AppActivate Title
    Sleep 400
    DoEvents
    Sleep 200
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        '去除多余的无效字符
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        '转换成数组
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------


Dim Myhwnd As Long
    '--------------
        Myhwnd = CLng(AllHwnd(20)) 'Uound(AllHwnd)=38,测试可知当t=20，对应Item ID
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '获取Item ID对应Rectangle坐标
        Call SetCursorPos(MyRect.Left + 60, MyRect.Top + 20)
        Call MouseClick(Title, MyRect.Left + 60, MyRect.Top + 20) '选中第一条Item NO
        Sleep 100
    '--------------

    '--------------
        Myhwnd = CLng(AllHwnd(2)) 'Uound(AllHwnd)=38,测试可知当t=2，对应Select
        PostMessage Myhwnd, BM_CLICK, 0, 0 '点击Select按钮.
            '*********
                '方式2: 按回车
                'PostMessage Myhwnd, WM_LBUTTONDOWN, MK_LBUTTON, 0
                'PostMessage Myhwnd, WM_LBUTTONUP, MK_LBUTTON, 0
            '*********
    '--------------


        Call PO_Warning '延时500ms
        
End Function
Private Function RIMSII_Sales_Order_Show()

Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim MyRect As RECT
Dim Title As String


    hwnd = FindWindow(vbNullString, "Sales Order Popup Search Window")
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Debug.Print "没有找到:Sales Order Popup Search Window"
            End
        End If
        hwnd = FindWindow(vbNullString, "Sales Order Popup Search Window")
    Loop
    
    Title = "Sales Order Popup Search Window"
    Set ws = CreateObject("WSCRIPT.SHELL")
    ws.AppActivate Title
    Delay 100
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        '去除多余的无效字符
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        '转换成数组
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------------------------
    

Dim Myhwnd As Long
    '--------------
        Myhwnd = CLng(AllHwnd(20)) 'Uound(AllHwnd)=38,测试可知当t=20，对应Item ID
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '获取Item ID对应Rectangle坐标
        Call SetCursorPos(MyRect.Left + 60, MyRect.Top + 20)
        Call MouseClick(Title, MyRect.Left + 60, MyRect.Top + 20) '选中第一条Item NO
        Delay 100
    '--------------
    
    '--------------
        Myhwnd = CLng(AllHwnd(2)) 'Uound(AllHwnd)=38,测试可知当t=2，对应Select
        PostMessage Myhwnd, BM_CLICK, 0, 0 '点击Select按钮.
            '*********
                '方式2: 按回车
                'PostMessage Myhwnd, WM_LBUTTONDOWN, MK_LBUTTON, 0
                'PostMessage Myhwnd, WM_LBUTTONUP, MK_LBUTTON, 0
            '*********
    '--------------

 
        Call RIMSII_Warning '延时500ms
        
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

Private Function RIMSII_Warning()

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


Private Function ClearClipboard()
'清空剪切板

Dim MyObj As New DataObject
    MyObj.SetText ""
    MyObj.PutInClipboard
    DoEvents
End Function

Private Function GetClipboard()
'获取剪切板内容

Dim MyObj As DataObject
    
    On Error Resume Next
    
    Set MyObj = New DataObject
    
    MyObj.GetFromClipboard
    GetClipboard = MyObj.GetText()
    DoEvents
    
End Function

Private Function SetClipboard(StrData As String)
Dim MyObj As New DataObject
    MyObj.SetText StrData
    MyObj.PutInClipboard
    DoEvents
End Function

Sub test_RIMSII_Comments()
        
        Call Init
        
    'SendMyUnicode Cells(12, 4).Text
 
'    hwnd = FindWindow(vbNullString, "Comments ")
    'RIMSII_Comments 12

'    Call ClearClipboard
'
'    Call SetClipboard(Cells(Trow, 4))
End Sub

Private Function RIMSII_Comments(Trow As Integer)
Dim hwnd As Long
Dim i As Integer
Dim t As Integer
Dim MyRect As RECT
Dim Title As String
Dim temp As String
Dim StrData As String
Dim Data() As String

        Delay 500

        
        '----------
        Call Init
        
        hwnd = setHwnd

    '-----------
        Delay 100
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
        Delay 50
        i = i + 1
    Loop
    
    
    'Print
    '------------
    Call GetWindowRect(hwnd, MyRect)
    Call SetCursorPos(MyRect.Left + 350, MyRect.Top + 48)
    Call MouseClick(Title, MyRect.Left + 60, MyRect.Top + 20)
    Delay 100
    '------------

        For i = 1 To 5
            SendMyData "Tab"
        Next i
        
        temp = Cells(Trow, 4).Text
        SendMyUnicode temp

        Delay 100
        SendMyData "%s"
        SendMyData "%c"

End Function


Private Sub test_RIMSII_DueDate()
   RIMSII_DueDate "2016-3-4"
End Sub

Private Function RIMSII_DueDate(DueDate As String)
Dim i As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String

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

Private Function RIMSII_New_Orders()
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


Private Sub Paste_ArrangeData()
    
    Sheets("PO Order form").Select
    Call PO_Paste_Data
    
    Sheets("RIMSII").Select
    ArrangeData
    
End Sub

Sub test()
    Date = "02/24/2016"
End Sub
Private Sub ArrangeData()

Dim temp As String
Dim Data() As String
 
 
 temp = [B5]
 
 If temp = "" Then
    ws_PopUp "请选择RBO!"
    Exit Sub
 End If
 
 Select Case temp
    
    Case "Carter's", "Maurices"
 
            temp = Sheets("PO Order form").Cells(1, 1) & Sheets("PO Order form").Cells(1, 2)
            
            If temp = "" Then
                Sheets("PO Order form").Select
                ws_PopUp "没有找到数据!"
                Exit Sub
            End If
            '============
                
                GetData 1, "RVL SO No"
                GetData 2, "Vendor"
                GetData 3, "Due Date"
                GetData 4, "Comment"
            
            '=============
        
            Cells(8, 2) = 12
            Cells(9, 2) = Cells(10000, 1).End(xlUp).Row
            
            
            Cells(8, 4) = "整理完成"
            
    Case "Pumpkin Patch"
        Call SplictSheetData
        
            Cells(8, 2) = 12
            Cells(9, 2) = Cells(10000, 1).End(xlUp).Row
            
            
            Cells(8, 4) = "整理完成"
    
End Select
    
End Sub
    
Private Function SplictSheetData()
'2016-02-26
Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim temp As String
Dim Data() As String
    
Dim r1 As Range
Dim SH As Worksheet
Dim MySH As Worksheet
    
    Set MySH = Sheets("RIMSII")
    Set SH = Sheets("PO Order form")
    
    'SH.Activate
    

Dim Endrow As Integer
Dim Srow As Integer
Dim Erow As Integer
Dim Trow As Integer
Dim Tcol As Integer
Dim MyCol As Integer
Dim myrow As Integer

Dim DueDate As String
Dim Vendor As String
Dim Comment As String
Dim CData() As String
Dim Sample As String
Dim StrData As String

    With SH
        '---------------
            'Vendor
            Set r1 = .Cells.Find("Vendor", , , xlPart)
            If Not r1 Is Nothing Then
                Data = VBA.Split(r1, ":")
                Vendor = "Avery Dennison (Guangzhou) Converted" '"Avery Dennison (Guangzhou) Converted Producted Ltd" 'Data(1)
                MyCol = r1.Column
            Else
                ws_PopUp "没有找到Vendor"
                Exit Function
            End If
            
            'DueDate
            Set r1 = .Cells.Find("Due Date", , , xlPart)
            If Not r1 Is Nothing Then
                Data = VBA.Split(r1, ":")
                DueDate = Data(1)
            Else
                ws_PopUp "没有找到Due Date"
                Exit Function
            End If
            
            'Comment
            Set r1 = .Cells.Find("comment", , , xlPart)
            If Not r1 Is Nothing Then
                Srow = r1.Row + 1
                t = 0
                For i = Srow To Srow + 10
                    temp = .Cells(i, MyCol)
                    temp = VBA.Trim(temp)
                    
                    If temp = "" Then
                        Exit For
                    End If
                    
                    ReDim Preserve CData(t)
                    temp = VBA.Replace(temp, VBA.ChrW(160), " ")
                    CData(t) = temp
                    t = t + 1
                Next i
            End If
        
        '---------------
        temp = "####################"
        Set r1 = .Cells.Find(temp, , , xlPart)
        If Not r1 Is Nothing Then
            myrow = r1.Row
            Tcol = r1.Column
            Erow = .Cells(50000, 1).End(xlUp).Row
            If Erow > myrow Then
                .Range(.Cells(myrow - 1, Tcol), .Cells(Erow, Tcol + 20)).ClearContents
            Else
                r1 = ""
            End If
        End If
        
        temp = "--------"
        Set r1 = .Cells.Find(temp, , , xlPart)
        If Not r1 Is Nothing Then
            MyCol = r1.Column
            myrow = r1.Row
        Else
            ws_PopUp "没有找到目标内容，请确认数据的正确性！"
            Exit Function
        End If
        
        Endrow = .Cells(60000, 1).End(xlUp).Row
        Trow = Endrow + 100
        .Cells(Trow, MyCol) = "####################"
        t = Trow + 1
        Srow = myrow + 1
        Erow = Endrow
        myrow = 12
        
        For i = Srow To Erow
            temp = .Cells(i, MyCol)
            temp = temp
            If InStr(1, temp, "---") Or VBA.Trim(temp) = "" Then
            Else
                
                Erase Data
                Call DataSplit(temp, Data)
                
                If UBound(Data) < 7 Then
                    ws_PopUp "Order form内容出错，请确保内容完整!"
                    End
                End If
                
                For j = 0 To UBound(Data)
                    .Cells(t, j + 1) = Data(j)
                Next j

                MySH.Cells(myrow, 1) = Data(4) 'RVL SO NO
                MySH.Cells(myrow, 2) = Vendor
                MySH.Cells(myrow, 3) = DueDate
                    
                    Comment = ""
                    If UBound(Data) = 8 Then
                        temp = ""
                        Sample = Data(8)
                        For j = 0 To UBound(CData)
                            temp = CData(j)
                            If InStr(1, VBA.LCase(temp), "phx so") Then
                                temp = temp & Data(1)
                            End If
                            
                            If InStr(1, VBA.LCase(temp), "sample") Then
                                k = InStr(1, temp, ":")
                                StrData = VBA.Mid(temp, 1, k)
                                temp = StrData & Sample
                            End If
                            
                            Comment = Comment & temp & VBA.Chr(10)
                            
                        Next j
                    Else
                        
                        
                        '--------------------
                        Sample = "NO"
                        For j = 0 To UBound(CData)
                            temp = CData(j)
                            If InStr(1, VBA.LCase(temp), "phx so") Then
                                temp = temp & Data(1)
                            End If
                            
                            If InStr(1, VBA.LCase(temp), "sample") Then
                                If InStr(1, VBA.LCase(temp), "s#") Then
                                Else
                                    k = InStr(1, temp, ":")
                                    StrData = VBA.Mid(temp, 1, k)
                                    temp = StrData & Sample
                                End If
                            End If
                            
                            Comment = Comment & temp & VBA.Chr(10)
                        Next j
                        '--------------------
                    End If
                MySH.Cells(myrow, 4) = Comment
                myrow = myrow + 1
                t = t + 1
            End If
        Next i
        
         '.Columns("B:Z").AutoFit

    End With
    
   
    
    
    
End Function

Private Function DataSplit(temp As String, Data() As String)
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As Integer

Dim tempData As String
Dim StrData As String
Dim Tdata As String
'Dim temp As String
'Dim Data() As String
 
    
    tempData = VBA.Space(2)
    
    '========================
    
    t = 0
    Tdata = SpecialCharData(temp)
    
    i = InStr(1, Tdata, tempData)
    Do
        
        DoEvents
        StrData = VBA.Mid(Tdata, 1, i - 1)
        ReDim Preserve Data(t)
        Data(t) = StrData
        t = t + 1
        Tdata = VBA.Trim(VBA.Mid(Tdata, i)) & tempData
        If Tdata = tempData Then
            Exit Do
        End If
        i = InStr(1, Tdata, tempData)
    Loop
 
End Function

Private Function SpecialCharData(Tdata As String) As String
Dim i As Integer
Dim StrData As String

    '---------特殊字符替换为Space(1)
    For i = 1 To VBA.Len(Tdata)
        StrData = VBA.Mid(Tdata, i, 1)
        If VBA.AscW(StrData) > 122 Then
            Tdata = VBA.Replace(Tdata, StrData, " ")
        End If
    Next i
    '---------
    
    SpecialCharData = Tdata
    
End Function
Private Function GetData(MyCol As Integer, temp As String, Optional AfterData As String)
Dim i As Integer
Dim j As Integer

Dim Srow As Integer
Dim Erow As Integer

Dim r1 As Range
Dim r2 As Range

Dim Flag As Integer

    With Sheets("PO Order form")
        
        Erow = .Cells(50000, 1).End(xlUp).Row
        
        If AfterData = "" Then
            Set r1 = .Cells.Find(temp, , , xlWhole)
            If Not r1 Is Nothing Then
                Srow = r1.Row + 1
                j = 12
                For i = Srow To Erow
                    Cells(j, MyCol) = .Cells(i, r1.Column).Value
                    j = j + 1
                Next i
            Else
                .Select
                ws_PopUp "没有找到：" & temp & "列！"
                Exit Function
            End If
        Else
            Flag = 0
            Set r2 = .Cells.Find(AfterData, , , xlPart)
            If Not r2 Is Nothing Then
                Set r1 = .Cells.Find(temp, r2, , xlWhole)
                If Not r1 Is Nothing Then
                    Srow = r1.Row + 1
                    j = 12
                    For i = Srow To Erow
                        Cells(j, MyCol) = .Cells(i, r1.Column).Value
                        j = j + 1
                    Next i
                Else
                    Flag = 1
                End If
            Else
                Flag = 1
            End If
            
            If Flag = 1 Then
                .Select
                ws_PopUp "没有找到：" & temp & "列！"
                End
            End If
        End If
        
        
    End With
            
End Function


Private Sub ClearData()
Dim Erow As Integer
    
    Erow = Cells(10000, 1).End(xlUp).Row
    If Erow > 11 Then
        Range(Cells(12, 1), Cells(Erow, 10)).ClearContents
        Range(Cells(12, 1), Cells(Erow, 10)).Interior.Color = xlNone
    End If
    
    Cells(8, 2) = ""
    Cells(9, 2) = ""
    Cells(8, 4) = ""
    '[B5] = ""
    
Dim DataSH As Worksheet

    Set DataSH = Sheets("PO Data")
    
    With DataSH
        Erow = DataSH.Cells(50000, 1).End(xlUp).Row
        
        If Erow > 1 Then
            .Range(.Cells(2, 1), .Cells(Erow, 10)).Interior.Color = xlNone
        End If
            
        .Cells.ClearContents
        
    End With
End Sub




Private Sub PO_ClearData()
Dim Erow As Integer
    
    Cells.ClearContents
    Cells.ClearFormats
    
    
    Rows("20:200").Delete
    
    'Erow = Cells(10000, 1).End(xlUp).Row
    'Range(Cells(1, 1), Cells(Erow, 20)).ClearContents
    'Range(Cells(1, 1), Cells(Erow, 20)).ClearFormats '
    'Range(Cells(1, 1), Cells(Erow, 20)).Interior.Color = xlNone
    
End Sub

Private Sub PO_Paste_Data()
    Range("A1").Select
    ActiveSheet.Paste
End Sub


Private Sub GetPO()
    Select Case [B5]
        Case "Pumpkin Patch"
            Call GetPO_Pumpkin_Patch
        Case "Carter's"
            Call GetPO_Carters
    End Select
End Sub

Private Sub GetPO_Pumpkin_Patch()

Dim WK As Workbook
Dim SH As Worksheet
Dim MySH As Worksheet
Dim Flag As Integer
Dim i As Integer
Dim j As Integer
Dim t As Integer
Dim r1 As Range
Dim Data() As String

    Set MySH = ActiveSheet
    
    Flag = 0
    For Each WK In Workbooks
        For Each SH In WK.Worksheets
            If SH.Cells(1, 12) = "po_no" Then
                Flag = 1
                Exit For
            End If
        Next
        If Flag = 1 Then
            Exit For
        End If
    Next
    
    If Flag = 0 Then
        ws_PopUp "没有找到report表！"
        Exit Sub
    End If
    
Dim Erow As Integer
Dim temp As String
Dim Trow As Integer
    
    Erow = Cells(60000, 1).End(xlUp).Row
    
    If Erow < 12 Then
        ws_PopUp "没有找到内容！"
        Exit Sub
    End If
    
    Flag = 0
    For i = 12 To Erow
        temp = Cells(i, 1)
        Set r1 = SH.Columns(1).Find(temp, , , xlWhole)
        If Not r1 Is Nothing Then
            Cells(i, 6) = SH.Cells(r1.Row, 12)
        Else
            Flag = 1
            Cells(i, 6).Interior.Color = vbRed
        End If
    Next i
    
    If Flag = 1 Then
        Cells(8, 4) = "存在SO#没有找到PO#"
    Else
        Cells(8, 4) = "提取PO#完成"
        Sheets("PO Data").Range(Sheets("PO Data").Cells(2, 1), Sheets("PO Data").Cells(Erow, 15)).Interior.Color = xlNone
    End If
    
    
    Dim TargetSH As Worksheet
    Dim DataSH As Worksheet
    
    Set TargetSH = Sheets("PO Order form")
    Set DataSH = Sheets("PO Data")
    
    DataSH.Cells.ClearContents
    
        temp = "####################"
        Set r1 = TargetSH.Cells.Find(temp, , , xlWhole)
        If r1 Is Nothing Then
            ws_PopUp "没有找到关键内容!"
            End
        Else
            Trow = r1.Row + 1
        End If
    
    With DataSH
        Erow = TargetSH.Cells(50000, 1).End(xlUp).Row
        
        temp = "Phx PO#,Rims PO#,Rims Item#,LLKK or LB,Qty"
        Erase Data
        Data = VBA.Split(temp, ",")
        For i = 1 To 5
            .Cells(1, i) = Data(i - 1)
        Next i
        
        t = 2
        For i = Trow To Erow
            .Cells(t, 1) = TargetSH.Cells(i, 2).Value
            .Cells(t, 3) = TargetSH.Cells(i, 6).Value
            .Cells(t, 5) = TargetSH.Cells(i, 7).Value
            t = t + 1
        Next i

        t = 12
        For i = 2 To Erow
            .Cells(i, 2) = MySH.Cells(t, 6)
            If MySH.Cells(t, 6).Interior.Color <> vbWhite Then
                .Cells(i, 2).Interior.Color = MySH.Cells(t, 6).Interior.Color
            End If
            t = t + 1
        Next i
    End With
    
    DataSH.Columns("A:Z").AutoFit
    
End Sub

Private Sub GetPO_Carters()
Dim WK As Workbook
Dim SH As Worksheet
Dim MySH As Worksheet
Dim Flag As Integer
Dim i As Integer
Dim j As Integer
Dim t As Integer
Dim r1 As Range

    Set MySH = ActiveSheet
    
    Flag = 0
    For Each WK In Workbooks
        For Each SH In WK.Worksheets
            If SH.Cells(1, 12) = "po_no" Then
                Flag = 1
                Exit For
            End If
        Next
        If Flag = 1 Then
            Exit For
        End If
    Next
    
    If Flag = 0 Then
        ws_PopUp "没有找到report表！"
        Exit Sub
    End If
    
Dim Erow As Integer
Dim temp As String
    
    Erow = Cells(60000, 1).End(xlUp).Row
    
    If Erow < 12 Then
        ws_PopUp "没有找到内容！"
        Exit Sub
    End If
    
    Flag = 0
    For i = 12 To Erow
        temp = Cells(i, 1)
        Set r1 = SH.Columns(1).Find(temp, , , xlWhole)
        If Not r1 Is Nothing Then
            Cells(i, 6) = SH.Cells(r1.Row, 12)
        Else
            Flag = 1
            Cells(i, 6).Interior.Color = vbRed
        End If
    Next i
    
    If Flag = 1 Then
        Cells(8, 4) = "存在SO#没有找到PO#"
    Else
        Cells(8, 4) = "提取PO#完成"
        Sheets("PO Data").Range(Sheets("PO Data").Cells(2, 1), Sheets("PO Data").Cells(Erow, 15)).Interior.Color = xlNone
    End If
    
    
    Dim TargetSH As Worksheet
    Dim DataSH As Worksheet
    
    Set TargetSH = Sheets("PO Order form")
    Set DataSH = Sheets("PO Data")
    
    DataSH.Cells.ClearContents
    
    
    With DataSH
        Erow = TargetSH.Cells(50000, 1).End(xlUp).Row
        
        For i = 1 To Erow
            For j = 1 To 6
                .Cells(i, j) = TargetSH.Cells(i, j).Value
            Next j
        Next i
        
        t = 11
        For i = 1 To Erow
            .Cells(i, 7) = MySH.Cells(t, 6)
            .Cells(i, 7).Interior.Color = MySH.Cells(t, 6).Interior.Color
            t = t + 1
        Next i
    End With
    
    DataSH.Columns("A:Z").AutoFit

End Sub
\
