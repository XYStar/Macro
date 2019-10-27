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

'��ȡdwExtraInfo
Declare Function GetMessageExtraInfo Lib "user32" () As Long
'�������ģ���˼����ж�  �������֧����Ļ���񣨽�ͼ����
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

'ö�ٴ����б��е����и����� (�����ͱ����д���)
'����ֵ Long�������ʾ�ɹ������ʾʧ��
'lpEnumFunc Long��ָ��Ϊÿ���Ӵ��ڶ����õ�һ��������ָ�롣��AddressOf�������ú����ڱ�׼ģʽ�µĵ�ַ
'lParam Long����ö���ڼ䣬���ݸ�dwcbkd32.ocx���ƿؼ�֮EnumWindows�¼���ֵ�����ֵ�ĺ������ɳ���Ա�涨��
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
'Ϊָ���ĸ�����ö���Ӵ���
'����ֵ Long�������ʾ�ɹ������ʾʧ��
'hWndParent Long����ö���Ӵ��ڵĸ����ڵľ��
'lpEnumFunc Long��Ϊÿ���Ӵ��ڵ��õĺ�����ָ�롣��AddressOf�������ú�����һ����׼ģ���еĵ�ַ
'lParam Long����ö���ڼ䣬���ݸ�dwcbkd32.ocx���ƿؼ�֮EnumWindows�¼���ֵ�����ֵ�ĺ������ɳ���Ա�涨�ġ���ԭ�ģ�Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.��
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'------------------------------------------------------------------
'Ѱ�Ҵ����б��е�һ������ָ�������Ķ������ڣ���vb��ʹ�ã�FindWindow�����һ����;�ǻ��ThunderRTMain������ش��ڵľ��������������������vbִ�г����һ���֡���þ���󣬿���api����GetWindowTextȡ��������ڵ����ƣ�����Ҳ��Ӧ�ó���ı��⣩
'lpClassName  String��ָ������˴��������Ŀ���ֹ��C���ԣ��ִ���ָ�룻����Ϊ�㣬��ʾ�����κ���
'lpWindowName  String��ָ������˴����ı������ǩ���Ŀ���ֹ��C���ԣ��ִ���ָ�룻����Ϊ�㣬��ʾ�����κδ��ڱ���
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'�ڴ����б���Ѱ����ָ����������ĵ�һ���Ӵ���
'hWnd1  Long�������в����ӵĸ����ڡ�����Ϊ�㣬��ʾʹ�����洰�ڣ�ͨ��˵�Ķ������ڶ�����Ϊ��������Ӵ��ڣ�����Ҳ������ǽ��в��ң�
'hWnd2  Long����������ں�ʼ���ҡ�����������ö�FindWindowEx�Ķ�ε����ҵ����������������Ӵ��ڡ�����Ϊ�㣬��ʾ�ӵ�һ���Ӵ��ڿ�ʼ����
'lpsz1  String�������������������ʾ����
'lpsz2  String�������������������ʾ����
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_GETTEXT = &HD

'----------------------------------------
'�ж�һ�������Ƿ�Ϊ��һ���ڵ��ӻ���������
Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
'�ж�ָ���ľ���Ƿ�Ϊһ���˵��ľ��
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
'�ж�һ�������Ƿ�Ϊ��
Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long

'�ж�һ�����ھ���Ƿ���Ч
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'�жϴ����Ƿ��ڻ״̬����vb��ʹ�ã����vb����Ϳؼ�������enabled���ԣ�
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

'----------------------------------------------------------------------------------------------
'���һ�����ڵľ�����ô�����ĳԴ�������ض��Ĺ�ϵ
'����ֵ  Long����wCmd������һ�����ڵľ������û���ҵ�������ڣ��������������򷵻���ֵ��������GetLastError
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'GetWindow()  wcmd
Private Const GW_HWNDFIRST = 0 'Ϊһ��Դ�Ӵ���Ѱ�ҵ�һ���ֵܣ�ͬ�������ڣ���Ѱ�ҵ�һ����������
Private Const GW_HWNDLAST = 1 'Ϊһ��Դ�Ӵ���Ѱ�����һ���ֵܣ�ͬ�������ڣ���Ѱ�����һ����������
Private Const GW_HWNDNEXT = 2 'ΪԴ����Ѱ����һ���ֵܴ���
Private Const GW_HWNDPREV = 3 'ΪԴ����Ѱ��ǰһ���ֵܴ���
Private Const GW_OWNER = 4 'Ѱ�Ҵ��ڵ�������
Private Const GW_CHILD = 5 'Ѱ��Դ���ڵĵ�һ���Ӵ���


Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'GetWindowDC ��ȡ�������ڣ������߿򡢹����������������˵��ȣ����豸����
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

'GetWindowExtEx ��ȡָ���豸�����Ĵ��ڷ�Χ
Declare Function GetWindowExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As Size) As Long

'GetWindowLong ��ָ�����ڵĽṹ��ȡ����Ϣ
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'GetWindowLong()
Private Const GWL_WNDPROC = (-4) '�ô��ڵĴ��ں����ĵ�ַ
Private Const GWL_HINSTANCE = (-6) 'ӵ�д��ڵ�ʵ���ľ��
Private Const GWL_HWNDPARENT = (-8) '�ô���֮���ľ������Ҫ��SetWindowWord���ı����ֵ
Private Const GWL_STYLE = (-16) '������ʽ
Private Const GWL_EXSTYLE = (-20) '��չ������ʽ
Private Const GWL_USERDATA = (-21) '������Ӧ�ó���涨
Private Const GWL_ID = (-12) '�Ի�����һ���Ӵ��ڵı�ʶ��

Private Const DWL_MSGRESULT = 0 '�ڶԻ������д����һ����Ϣ���ص�ֵ
Private Const DWL_DLGPROC = 4 '������ڵĶԻ�������ַ
Private Const DWL_USER = 8 '������Ӧ�ó���涨


'GetWindowText ȡ��һ������ı��⣨caption�����֣�����һ���ؼ������ݣ���vb��ʹ�ã�ʹ��vb�����ؼ���caption��text���ԣ�
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'GetWindowsDirectory ��������ܻ�ȡWindowsĿ¼������·�����������Ŀ¼������˴����windowsӦ�ó����ļ�����ʼ���ļ�
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'GetWindowTextLength ���鴰�ڱ������ֻ�ؼ����ݵĳ��̣���vb��ʹ�ã�ֱ��ʹ��vb�����ؼ���caption��text���ԣ�
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'GetWindowOrgEx ��ȡָ���豸�������߼����ڵ����
Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lppoint As POINTAPI) As Long

'GetWindowThreadProcessId ��ȡ��ָ�����ڹ�����һ���һ�����̺��̱߳�ʶ��
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


'----------------------------------------------------------------------------------------------

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'==============================          GetMenu              =======================================

'GetMenu
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'hwnd Long�����ڵľ��
'bRevert Long������ΪTRUE����ʾ����ԭʼ��ϵͳ�˵�
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'hMenu Long���˵��ľ��
'nPos Long����Ŀ�ڲ˵��е�λ�á���һ����Ŀ�ı��Ϊ0
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

'hMenu Long���˵����
'wIDItem Long�������յĲ˵���Ŀ�ı�ʶ���������wFlags������������MF_BYCOMMAND��־���������������ָ��Ҫ�ı�Ĳ˵���Ŀ������ID��������õ���MF_BYPOSITION��־���������������ָ����Ŀ�ڲ˵��е�λ�ã���һ����Ŀ��λ��Ϊ0��
'lpString String��ָ��һ��Ԥ�ȶ���õ��ִ����������Ա�Ϊ�˵���Ŀװ���ִ�
'nMaxCount Long������lpString�������е�����ַ�����+1
'wFlag Long������MF_BYCOMMAND��MF_BYPOSITION��ȡ����wID����������

Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long


Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&

'hMenu Long���˵��ľ��
'un Long���˵���Ŀ�Ĳ˵�ID��λ��
'b Boolean����unָ��������Ŀλ�ã���ΪTRUE����ָ������һ���˵�ID����ΪFALSE
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


'��һ��������װ��ָ���˵���Ŀ����Ļ������Ϣ
'����ֵ Long��TRUE�����㣩��ʾ�ɹ������򷵻��㡣������GetLastError
'hWnd Long������ָ���˵��򵯳�ʽ�˵���һ�����ڵľ��
'hMenu Long���˵��ľ��
'uItem Long�������Ĳ˵���Ŀ��λ�û�˵�ID
'lprcItem RECT��������ṹ��װ�ز˵���Ŀ��λ�ü���С��������Ļ�����ʾ��
Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
'��ȡ�ؼ�����
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Public Const WM_CHAR = &H102
Public ChildHwnd As String '�����Ӵ���ؼ����


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


'����һ���ص�����, �������ģ����. ��������ָ�����ڵ��Ӵ���(�ؼ�). ��������е� hWnd ��Ϊ�Ӵ���(�ؼ�)���
Public Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
ChildHwnd = ChildHwnd & "," & hwnd
EnumChildProc = 1
End Function


' ����: FGetClassName
' ����: ����ָ�������е�����
' ����: hWnd ָ�����ڵľ��
' ����: ָ�����ڵ�����

Public Function FGetClassName(hwnd As Long) As String
Dim ClassName As String
Dim Ret As Long
'��仺��(������̫С���ػ᲻����).
 ClassName = Space(256)

 '���� GetClassName ����, ����ֵΪ��������ʵ�ʳ���.
 Ret = GetClassName(hwnd, ClassName, 256)

 '������������. Ret Ϊ��һ�����õ�����������ʵ�ʳ���.
 FGetClassName = Left(ClassName, Ret)
End Function

' ����: GetText
' ����: ����ָ������(���ı���)�е�����
' ����: WindowHandle ָ�����ڵľ��
' ����: ָ�����ڵ�����
Public Function GetText(WindowHandle As Long) As String
Dim strBuffer As String '�ַ�������
Dim Char As String '�������������Դ��ָ�

    '��仺��(������̫С���ػ᲻����).
    strBuffer = Space(255)

    '������Ϣ EM_GETPASSWORDCHAR(������������) ��ָ������. ���ﷵ�������Char(������� Char=*).
    Char = SendMessage(WindowHandle, &HD2, 0, 0)

    '������Ϣ EM_SETPASSWORDCHAR(������������) ��ָ������. ����������0(Null), ����ȥ����������.
    PostMessage WindowHandle, &HCC, 0, 0

    '�����Edit�ؼ���ȴ���Ϣ���ͳɹ�, ���ȴ����뱻ȥ��.
    If InStr("Edit", FGetClassName(WindowHandle)) And Char <> "0" Then Sleep (10)

    '������Ϣ WM_GETTEXT(������������) ��ָ������. ����õ�Edit�ؼ�������, ������. ע��"ByVal", ����������VB����.
    SendMessage WindowHandle, &HD, 255, ByVal strBuffer

    '������Ϣ EM_SETPASSWORDCHAR(������������) ��ָ������. ��������ΪChar, ���ָ�ԭ������.
    PostMessage WindowHandle, &HCC, ByVal Char, 0

    '����������������(����), ֮����Ҫ��Trimȥ�ո�����Ϊ��һ�����ÿո������255���ַ�.
    GetText = Trim(strBuffer)
End Function


''Ϊָ���ĸ�����ö���Ӵ���
''����ֵ Long�������ʾ�ɹ������ʾʧ��
''hWndParent Long����ö���Ӵ��ڵĸ����ڵľ��
''lpEnumFunc Long��Ϊÿ���Ӵ��ڵ��õĺ�����ָ�롣��AddressOf�������ú�����һ����׼ģ���еĵ�ַ
''lParam Long����ö���ڼ䣬���ݸ�dwcbkd32.ocx���ƿؼ�֮EnumWindows�¼���ֵ�����ֵ�ĺ������ɳ���Ա�涨�ġ���ԭ�ģ�Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.��
'Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal LPARAM As Long) As Long

'���������������ÿ�ε��ö���õ���һ���Ӵ��壨�ؼ����ľ��������ֵ��hWnd,ʵ��ʹ���У��Ұ������Ӿ�������ChildHwnd�ַ����У�������ϣ���

'Dim AllHwnd() As String

'ȥ���������Ч�ַ�
'ChildHwnd =vba. Mid(ChildHwnd, 2)
'ת��������
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
        'ֹͣ���Һ�������0
        EnumWindowsProc = 0
    Else
        '�������Һ�������1
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
'y0 = y0 + 326 'Y+316-->Y+336 ,ȡ�м�ֵ:326

'Call SetCursorPos(x0, y0)
'MouseClick temp, x0, y0
'Call SetCursorPos(x1, y1)
'
ChildHwnd = ""

'���� EnumChidWindows ������ʼ����ָ�����ڵ��Ӵ���(�ؼ�). ��һ��������ָ�����ڵľ��, �ڶ�������Ϊ����ص������ĵ�ַ(��AddressOf���������), �������������ù�...
t = EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)

'ȥ���������Ч�ַ�
ChildHwnd = VBA.Mid(ChildHwnd, 2)
'ת��������
AllHwnd = VBA.Split(ChildHwnd, ",")

Data = ""

Dim Myhwnd As Long


'Button ��ť�ؼ�
'COMBOBOX ��Ͽ�ؼ�
'Edit �༭��ؼ�
'ListBox �б��ؼ�
'ScrollBar �������ؼ�
'Static ��̬�ؼ�

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
'                        '��ʽ1:
'                        'PostMessage Myhwnd, BM_CLICK, 0, 0
'
'                        '��ʽ2:
'                        'PostMessage Myhwnd, WM_LBUTTONDOWN, MK_LBUTTON, 0
'                        'PostMessage Myhwnd, WM_LBUTTONUP, MK_LBUTTON, 0
'
'        '----------------------------------
'                        '��ʽ3:
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
'            If t = 20 Then 'Item ID,���Կ�֪��t=20����ӦItem ID
'                Call GetWindowRect(Myhwnd, myRect) '��ȡItem ID��ӦRectangle����
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

'���� EnumChidWindows ������ʼ����ָ�����ڵ��Ӵ���(�ؼ�). ��һ��������ָ�����ڵľ��, �ڶ�������Ϊ����ص������ĵ�ַ(��AddressOf���������), �������������ù�...
t = EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)

'ȥ���������Ч�ַ�
ChildHwnd = VBA.Mid(ChildHwnd, 2)
'ת��������
AllHwnd = VBA.Split(ChildHwnd, ",")

Data = ""

Dim Myhwnd As Long


'Button ��ť�ؼ�
'COMBOBOX ��Ͽ�ؼ�
'Edit �༭��ؼ�
'ListBox �б��ؼ�
'ScrollBar �������ؼ�
'Static ��̬�ؼ�

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
                    PostMessage EditHwnd, WM_CHAR, Asc(StrData), 0   ' ����һ�� A �ַ�
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
''                    PostMessage Myhwnd, WM_CHAR, Asc(Strdata), 0   ' ����һ�� A �ַ�
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
        MsgBox "û���ҵ�:" & Title
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

