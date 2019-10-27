Attribute VB_Name = "A_Function_For_RIMSII_SO"
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

'��ȡdwExtraInfo
Private Declare Function GetMessageExtraInfo Lib "user32" () As Long
'�������ģ���˼����ж�  �������֧����Ļ���񣨽�ͼ����
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

'ö�ٴ����б��е����и����� (�����ͱ����д���)
'����ֵ Long�������ʾ�ɹ������ʾʧ��
'lpEnumFunc Long��ָ��Ϊÿ���Ӵ��ڶ����õ�һ��������ָ�롣��AddressOf�������ú����ڱ�׼ģʽ�µĵ�ַ
'lParam Long����ö���ڼ䣬���ݸ�dwcbkd32.ocx���ƿؼ�֮EnumWindows�¼���ֵ�����ֵ�ĺ������ɳ���Ա�涨��
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
'Ϊָ���ĸ�����ö���Ӵ���
'����ֵ Long�������ʾ�ɹ������ʾʧ��
'hWndParent Long����ö���Ӵ��ڵĸ����ڵľ��
'lpEnumFunc Long��Ϊÿ���Ӵ��ڵ��õĺ�����ָ�롣��AddressOf�������ú�����һ����׼ģ���еĵ�ַ
'lParam Long����ö���ڼ䣬���ݸ�dwcbkd32.ocx���ƿؼ�֮EnumWindows�¼���ֵ�����ֵ�ĺ������ɳ���Ա�涨�ġ���ԭ�ģ�Value that is passed to the EnumWindows event of the dwcbkd32.ocx custom control during enumeration. The meaning of this value is defined by the programmer.��
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long


Private Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'------------------------------------------------------------------
'Ѱ�Ҵ����б��е�һ������ָ�������Ķ������ڣ���vb��ʹ�ã�FindWindow�����һ����;�ǻ��ThunderRTMain������ش��ڵľ��������������������vbִ�г����һ���֡���þ���󣬿���api����GetWindowTextȡ��������ڵ����ƣ�����Ҳ��Ӧ�ó���ı��⣩
'lpClassName  String��ָ������˴��������Ŀ���ֹ��C���ԣ��ִ���ָ�룻����Ϊ�㣬��ʾ�����κ���
'lpWindowName  String��ָ������˴����ı������ǩ���Ŀ���ֹ��C���ԣ��ִ���ָ�룻����Ϊ�㣬��ʾ�����κδ��ڱ���
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'�ڴ����б���Ѱ����ָ����������ĵ�һ���Ӵ���
'hWnd1  Long�������в����ӵĸ����ڡ�����Ϊ�㣬��ʾʹ�����洰�ڣ�ͨ��˵�Ķ������ڶ�����Ϊ��������Ӵ��ڣ�����Ҳ������ǽ��в��ң�
'hWnd2  Long����������ں�ʼ���ҡ�����������ö�FindWindowEx�Ķ�ε����ҵ����������������Ӵ��ڡ�����Ϊ�㣬��ʾ�ӵ�һ���Ӵ��ڿ�ʼ����
'lpsz1  String�������������������ʾ����
'lpsz2  String�������������������ʾ����
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

'----------------------------------------
'�ж�һ�������Ƿ�Ϊ��һ���ڵ��ӻ���������
Private Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
'�ж�ָ���ľ���Ƿ�Ϊһ���˵��ľ��
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
'�ж�һ�������Ƿ�Ϊ��
Private Declare Function IsRectEmpty Lib "user32" (lpRect As RECT) As Long

'�ж�һ�����ھ���Ƿ���Ч
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
'�жϴ����Ƿ��ڻ״̬����vb��ʹ�ã����vb����Ϳؼ�������enabled���ԣ�
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long

'----------------------------------------------------------------------------------------------
'���һ�����ڵľ�����ô�����ĳԴ�������ض��Ĺ�ϵ
'����ֵ  Long����wCmd������һ�����ڵľ������û���ҵ�������ڣ��������������򷵻���ֵ��������GetLastError
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

'GetWindow()  wcmd
Private Const GW_HWNDFIRST = 0 'Ϊһ��Դ�Ӵ���Ѱ�ҵ�һ���ֵܣ�ͬ�������ڣ���Ѱ�ҵ�һ����������
Private Const GW_HWNDLAST = 1 'Ϊһ��Դ�Ӵ���Ѱ�����һ���ֵܣ�ͬ�������ڣ���Ѱ�����һ����������
Private Const GW_HWNDNEXT = 2 'ΪԴ����Ѱ����һ���ֵܴ���
Private Const GW_HWNDPREV = 3 'ΪԴ����Ѱ��ǰһ���ֵܴ���
Private Const GW_OWNER = 4 'Ѱ�Ҵ��ڵ�������
Private Const GW_CHILD = 5 'Ѱ��Դ���ڵĵ�һ���Ӵ���


Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

'GetWindowDC ��ȡ�������ڣ������߿򡢹����������������˵��ȣ����豸����
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

'GetWindowExtEx ��ȡָ���豸�����Ĵ��ڷ�Χ
Private Declare Function GetWindowExtEx Lib "gdi32" (ByVal hDC As Long, lpSize As Size) As Long

'GetWindowLong ��ָ�����ڵĽṹ��ȡ����Ϣ
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

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
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'GetWindowsDirectory ��������ܻ�ȡWindowsĿ¼������·�����������Ŀ¼������˴����windowsӦ�ó����ļ�����ʼ���ļ�
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'GetWindowTextLength ���鴰�ڱ������ֻ�ؼ����ݵĳ��̣���vb��ʹ�ã�ֱ��ʹ��vb�����ؼ���caption��text���ԣ�
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

'GetWindowOrgEx ��ȡָ���豸�������߼����ڵ����
Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lppoint As POINTAPI) As Long

'GetWindowThreadProcessId ��ȡ��ָ�����ڹ�����һ���һ�����̺��̱߳�ʶ��
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long


'----------------------------------------------------------------------------------------------

Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


'==============================          GetMenu              =======================================
 
'GetMenu
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'hwnd Long�����ڵľ��
'bRevert Long������ΪTRUE����ʾ����ԭʼ��ϵͳ�˵�
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
'hMenu Long���˵��ľ��
'nPos Long����Ŀ�ڲ˵��е�λ�á���һ����Ŀ�ı��Ϊ0
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

'hMenu Long���˵����
'wIDItem Long�������յĲ˵���Ŀ�ı�ʶ���������wFlags������������MF_BYCOMMAND��־���������������ָ��Ҫ�ı�Ĳ˵���Ŀ������ID��������õ���MF_BYPOSITION��־���������������ָ����Ŀ�ڲ˵��е�λ�ã���һ����Ŀ��λ��Ϊ0��
'lpString String��ָ��һ��Ԥ�ȶ���õ��ִ����������Ա�Ϊ�˵���Ŀװ���ִ�
'nMaxCount Long������lpString�������е�����ַ�����+1
'wFlag Long������MF_BYCOMMAND��MF_BYPOSITION��ȡ����wID����������

Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long


Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&

'hMenu Long���˵��ľ��
'un Long���˵���Ŀ�Ĳ˵�ID��λ��
'b Boolean����unָ��������Ŀλ�ã���ΪTRUE����ָ������һ���˵�ID����ΪFALSE
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


'��һ��������װ��ָ���˵���Ŀ����Ļ������Ϣ
'����ֵ Long��TRUE�����㣩��ʾ�ɹ������򷵻��㡣������GetLastError
'hWnd Long������ָ���˵��򵯳�ʽ�˵���һ�����ڵľ��
'hMenu Long���˵��ľ��
'uItem Long�������Ĳ˵���Ŀ��λ�û�˵�ID
'lprcItem RECT��������ṹ��װ�ز˵���Ŀ��λ�ü���С��������Ļ�����ʾ��
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
'��ȡ�ؼ�����
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102

Private ChildHwnd As String   '�����Ӵ���ؼ����


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

'---------
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
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

'---------


Dim XY As POINTAPI
Dim AppData As String
'===========2016============


'����һ���ص�����, �������ģ����. ��������ָ�����ڵ��Ӵ���(�ؼ�). ��������е� hWnd ��Ϊ�Ӵ���(�ؼ�)���
Private Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
ChildHwnd = ChildHwnd & "," & hwnd
EnumChildProc = 1
Debug.Print EnumChildProc = 1
End Function


' ����: FGetClassName
' ����: ����ָ�������е�����
' ����: hWnd ָ�����ڵľ��
' ����: ָ�����ڵ�����

Private Function FGetClassName(hwnd As Long) As String
Dim ClassName As String
Dim Ret As Long
'��仺��(������̫С���ػ᲻����).
 ClassName = VBA.Space(256)

 '���� GetClassName ����, ����ֵΪ��������ʵ�ʳ���.
 Ret = GetClassName(hwnd, ClassName, 256)
 
 '������������. Ret Ϊ��һ�����õ�����������ʵ�ʳ���.
 FGetClassName = VBA.Left(ClassName, Ret)
End Function

' ����: GetText
' ����: ����ָ������(���ı���)�е�����
' ����: WindowHandle ָ�����ڵľ��
' ����: ָ�����ڵ�����
Private Function GetText(WindowHandle As Long) As String
Dim strBuffer As String '�ַ�������
Dim Char As String '�������������Դ��ָ�

    '��仺��(������̫С���ػ᲻����).
    strBuffer = VBA.Space(255)
    '������Ϣ EM_GETPASSWORDCHAR(������������) ��ָ������. ���ﷵ�������Char(������� Char=*).
    Char = SendMessage(WindowHandle, &HD2, 0, 0)
    '������Ϣ EM_SETPASSWORDCHAR(������������) ��ָ������. ����������0(Null), ����ȥ����������.
    PostMessage WindowHandle, &HCC, 0, 0
    '�����Edit�ؼ���ȴ���Ϣ���ͳɹ�, ���ȴ����뱻ȥ��.
    If InStr("Edit", FGetClassName(WindowHandle)) And Char <> "0" Then Sleep (10)
    '������Ϣ WM_GETTEXT(������������) ��ָ������. ����õ�Edit�ؼ�������, ������. ע��"ByVal".
    SendMessage WindowHandle, &HD, 255, ByVal strBuffer
    '������Ϣ EM_SETPASSWORDCHAR(������������) ��ָ������. ��������ΪChar, ���ָ�ԭ������.
    PostMessage WindowHandle, &HCC, ByVal Char, 0
    '����������������(����), ֮����Ҫ��Trimȥ�ո�����Ϊ��һ�����ÿո������255���ַ�.
    GetText = VBA.Replace(VBA.Trim(strBuffer), VBA.Chr(0), "")
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

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

    Dim lpBuffer As String * 1024
    Dim dwWindowCaption As String
    Dim lpLength As Long
    'GetWindowText ȡ��һ������ı��⣨caption�����֣�����һ���ؼ������ݣ���vb��ʹ�ã�ʹ��vb�����ؼ���caption��text���ԣ�
    lpLength = GetWindowText(hwnd, lpBuffer, 1024)
    dwWindowCaption = VBA.Left(lpBuffer, lpLength)
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

Private Sub test_Step_F3()
    Step_F3 "2872", "4556/6", "Test_By_Star"
End Sub

Public Function Step_F3(Master As String, _
                        Billto As String, _
                        PO As String, _
                        OrderRecdDate As String, _
                        Optional returnOrderNO As String, _
                        Optional myrow As Integer, _
                        Optional Col As Dictionary)

Dim Data() As String
Dim MyData As String
Dim temp As String

    If GetESC Then SetESC
        
    'F3
    '======
        'Master Customer
        '-------
            If Master = "" Then
                Exit Function
            End If
            
            If InStr(Master, "/") Then
                Data = VBA.Split(Master, "/")
                MyData = Data(0) ' "1655"
                Master_Customer MyData '���롰1655��
                MyData = Master '"ѡ�� 1655/6"
                Popup_Search MyData
            Else
                MyData = Master
                Master_Customer MyData
            End If
            
            If Master <> "" And myrow <> 0 Then
                Cells(myrow, Col("master customer")).Interior.Color = vbGreen
            End If
            
        '-------
        'Bill to Customer
        '-------
            If InStr(Billto, "/") Then
                Data = VBA.Split(Billto, "/")
                MyData = Data(0) ' "4596"
                Bill_to_Customer MyData
                MyData = Billto ' "4596/9"
                Popup_Search MyData
            Else
                MyData = Billto
                If MyData <> "" Then
                    Bill_to_Customer MyData
                End If
            End If
        
            If Billto <> "" And myrow <> 0 Then
                Cells(myrow, Col("bill to customer")).Interior.Color = vbGreen
            End If
            
        '-------
        'Customer PO NO
        '-------
            MyData = PO
            If MyData <> "" Then
                Customer_PO_NO MyData
            End If
            If PO <> "" And myrow <> 0 Then
                Cells(myrow, Col("customer po no.")).Interior.Color = vbGreen
            End If
        '-------
        
        'Order_Recd_Date
        '-------
            MyData = OrderRecdDate
            If MyData <> "" Then
                Call Order_Recd_Date(MyData)
            End If
            If OrderRecdDate <> "" And myrow <> 0 Then
                Cells(myrow, Col("Order Recd Date")).Interior.Color = vbGreen
            End If
        '-------
        
            Save_Form
        
        '��ȡItem NO.
        '------
            returnOrderNO = Get_OrderNO
            Delay 300
        '------
    '======
    
End Function

Public Function Step_F4(Shipto As String, _
                        Optional myrow As Integer, _
                        Optional Col As Dictionary)
Dim Data() As String
Dim MyData As String
Dim temp As String

    If GetESC Then SetESC
    
    'F4
    '======
        '-------
            If Shipto = "" Then
                Exit Function
            End If
            
            If InStr(Shipto, "/") Then
                Data = VBA.Split(Shipto, "/")
                MyData = Data(0) '"4596"
                Ship_To_Address MyData
                MyData = Shipto '"4596/11"
                Popup_Search MyData
            Else
                MyData = Shipto
                If MyData <> "" Then
                    Call Ship_To_Address(Shipto)
                End If
            End If

            If Shipto <> "" And myrow <> 0 Then
                Cells(myrow, Col("ship to address")).Interior.Color = vbGreen
            End If
            
            Save_Form
        '-------
    '======
End Function


Public Function Step_F5(ItemNO As String, _
                        Total As String, _
                        Price As String, _
                        Optional StyleNO As String, _
                        Optional myrow As Integer, _
                        Optional Col As Dictionary)

    If GetESC Then SetESC
    
    'F5
    '======
        '-------
            If ItemNO = "" Then
                Exit Function
            End If
            
            Call Order_Lines("Buy/Sell", ItemNO, Total, Price, StyleNO, myrow, Col)
            Save_Form
        '-------
    '======
End Function

Public Function Step_F7(Dict As Dictionary, _
                        Optional MP As Dictionary)

    If GetESC Then SetESC
    
    'F7
    '======
        '-------
            If Dict Is Nothing Then
                Exit Function
            End If
            
            Call Add_Size_QTY(Dict, MP)
            Save_Form
        '-------
    '======
End Function

Public Function Step_F8(ReqDate As String, _
                        PromisedDate As String, _
                        Optional myrow As Integer, _
                        Optional Col As Dictionary)

    If GetESC Then SetESC
    
    'F8
    '======
        '-------
            If ReqDate = "" Or PromisedDate = "" Then
                Exit Function
            End If
            
            Call Schedule(ReqDate, PromisedDate, myrow, Col)

            Save_Form
        '-------
    '======
    
End Function


Private Sub test_Set_Dict()
    Dim Dict As Dictionary
    
    Set Dict = Set_Dict("XL|XXS", "100|99")
    Set Dict = Set_Dict("XS", "100", Dict)
    
    For i = 0 To Dict.Count - 1
        Debug.Print Dict.Keys(i)
        Debug.Print Dict.Items(i)
    Next i
    
End Sub


Public Function Set_Dict(Size As String, QTY As String, Optional MyDict As Dictionary) As Dictionary
    Dim Dict As Dictionary
    Dim temp As String
    Dim i As Integer
    Dim Data() As String
    Dim Tdata() As String
        
        If MyDict Is Nothing Then
            Set Dict = CreateObject("Scripting.Dictionary")
        Else
            Set Dict = MyDict
        End If
        
        If InStr(Size, "//") Then
            Erase Data
            Erase Tdata
            Data = VBA.Split(Size, "//")
            Tdata = VBA.Split(QTY, "//")
            For i = 0 To UBound(Data)
                Dict.Add Data(i), Tdata(i)
            Next i
        Else
            Dict.Add Size, QTY
        End If
            
        Set Set_Dict = Dict
        
End Function

Private Function test_For_Control_Sales_Order()
Dim hwnd As Long
Dim AllHwnd() As String

Dim i As Integer
Dim t As Long
Dim EditHwnd As Long
Dim temp As String
Dim StrData As String
Dim Myhwnd As Long
    
        Call Init
        
        hwnd = setHwnd

'        '------------
'        Dim TestFlag As Integer
'        TestFlag = 1
'        If TestFlag = 1 Then
'            Dim Title As String
'
'            Title = "Sales Order Popup Search Window"
'            hwnd = FindWindow(vbNullString, Title)
'            t = 0
'            Do While hwnd = 0
'                DoEvents
'                If GetESC Then
'                    Debug.Print "��ESCֹͣ����,����Ϊû���ҵ�:" & Title
'                    End
'                End If
'                hwnd = FindWindow(vbNullString, Title)
'                Delay 100
'                t = t + 1
'                If t > 15 Then 'ѭ��1.5��
'                    Debug.Print "û�д�:" & Title
'                    Exit Function
'                End If
'            Loop
'
'
'            Set ws = CreateObject("WSCRIPT.SHELL")
'            ws.AppActivate Title
'        End If
'        '------------
        
        '-------------------
            ChildHwnd = ""
            Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
            ChildHwnd = VBA.Mid(ChildHwnd, 2)
            AllHwnd = VBA.Split(ChildHwnd, ",")
        '---------------------

        '0722 Master Customer
        '-----------
        '��ʼ��Sales Order View ��Index��New �򿪵�Sales Order View��Index�ǲ�һ����
'            Dim MasterHwndIndex As Integer
'            Myhwnd = CLng(AllHwnd(49))
'            temp = FGetClassName(Myhwnd)
'            If VBA.LCase(temp) = "edit" Then
'                MasterHwndIndex = 49
'            Else
'                MasterHwndIndex = 44
'            End If
            '---------����Master Customer
'                Myhwnd = CLng(AllHwnd(MasterHwndIndex))
'                temp = "1655"
'                For i = 1 To Len(temp)
'                    Strdata = Mid(temp, i, 1)
'                    PostMessage Myhwnd, WM_CHAR, Asc(Strdata), 0
'                Next i
'                Sleep 100
'                PostMessage Myhwnd, WM_KEYDOWN, VK_ENTER, 0
'                PostMessage Myhwnd, WM_KEYUP, VK_ENTER, 0
            '---------
       '-----------



Dim lpBuffer As String * 256
Dim dwWindowCaption As String
Dim lpLength As Long

        Debug.Print UBound(AllHwnd)

        For t = 0 To UBound(AllHwnd)
            
            Myhwnd = CLng(AllHwnd(t))
            temp = FGetClassName(Myhwnd)
            
            
            
            '--------
'                lpBuffer = ""
'                lpLength = GetWindowText(Myhwnd, lpBuffer, 255)
'                If lpLength <> 0 Then
'                    dwWindowCaption = VBA.Left(lpBuffer, lpLength)
'                Else
'                    dwWindowCaption = ""
'                End If
            '--------
            'Debug.Print t & ": " & temp
            'Debug.Print GetText(Myhwnd)
            
            'Debug.Print GetText(CLng(AllHwnd(129)))
            
            If InStr(VBA.LCase(GetText(Myhwnd)), "330") Then
                'Strdata = "test"
                'PostMessage Myhwnd, WM_CHAR, Asc(Strdata), 0
                t = t '120

            End If
            
            Debug.Print GetText(Myhwnd)
            
            If t = 129 Then
                Debug.Print GetText(Myhwnd)
            End If
            
            If t > 66 Then '68: Order Line 1 of 1
            
                If VBA.Len(VBA.Trim(GetText(Myhwnd))) > 1 Then
                    
                    i = i
                End If
                
            End If
            
            
'            If t = 49 Then
'                Debug.Print GetText(Myhwnd)
'            End If
            
 
            If t = 44 Then
                t = t
            End If
            
            Select Case LCase(temp)
                Case "button"
                    
                    
                    
'                    If dwWindowCaption = "&Reset" Then
'                        temp = OrderNO
'                        For i = 1 To Len(temp)
'                            Strdata = Mid(temp, i, 1)
'                            PostMessage EditHwnd, WM_CHAR, Asc(Strdata), 0
'                        Next i
'
'                        hwnd = GetParent(EditHwnd)
'
'                        Sleep 100
'                        PostMessage hwnd, WM_KEYDOWN, VK_ENTER, 0
'                        PostMessage hwnd, WM_KEYUP, VK_ENTER, 0
'
'                        Delay 500
'
'                        Exit For
'                    End If
                    
                Case "edit"
                    
                    EditHwnd = Myhwnd
'                    If t = 83 Then 'Order Type
'                        'SendStrToSys EditHwnd, "GAP-GKLBL-198585-G-CN"
'                        PostMessage EditHwnd, WM_KEYDOWN, VK_ENTER, 0
'                        PostMessage EditHwnd, WM_KEYUP, VK_ENTER, 0
'                    End If
                    
                    
                    
'                    If t = MasterHwndIndex Then 'Master Customer
                        temp = "A"
                        For i = 1 To Len(temp)
                            StrData = Mid(temp, i, 1)
                            PostMessage EditHwnd, WM_CHAR, Asc(StrData), 0
                        Next i
''                        Sleep 100
''                        PostMessage EditHwnd, WM_KEYDOWN, VK_ENTER, 0
''                        PostMessage EditHwnd, WM_KEYUP, VK_ENTER, 0
''
''                        t = t
'
'                    End If
'
'                    If t > 27 And t <> 44 Then
''                        temp = "4596"
''                        For i = 1 To Len(temp)
''                            Strdata = Mid(temp, i, 1)
''                            PostMessage EditHwnd, WM_CHAR, Asc(Strdata), 0
''                        Next i
''                        Sleep 100
''                        PostMessage EditHwnd, WM_KEYDOWN, VK_ENTER, 0
''                        PostMessage EditHwnd, WM_KEYUP, VK_ENTER, 0
''
''                        t = t
'                    End If
                Case Else
'                    If t = 84 Then 'Order Type
'                        SendStrToSys Myhwnd, "Buy/Sell"
'                    End If
                    
            End Select
        
        Next t
End Function

Public Sub test_Show()
    test_Sales_Order
    SendMyData "^n"
    Delay 600
    test_Sales_Order
    SendMyData "^n"
    Delay 600
End Sub

Private Sub test_Sales_Order()

Dim MyData As String
    
    If GetESC Then SetESC
    
    'F3
    '======
        'Master Customer
        '-------
            MyData = "1655"
            Master_Customer MyData
            MyData = "1655/6"
            Popup_Search MyData
        '-------
        'Bill to Customer
        '-------
            MyData = "4596"
            Bill_to_Customer MyData
            MyData = "4596/9"
            Popup_Search MyData
        '-------
        'Customer PO NO
        '-------
            MyData = "Test_By_Star"
            Customer_PO_NO MyData
        '-------
            Save_Form
            
        '-----
            Debug.Print "Order NO.: " & Get_OrderNO
            
    '======
    
    'F4
    '======
        '-------
            MyData = "4596"
            Ship_To_Address MyData
            MyData = "4596/11"
            Popup_Search MyData
            Save_Form
        '-------
    '======
    
    'F5
    '======
        '-------
            Order_Lines "Buy/Sell", "GAP-GKLBL-198585-G-CN", "1900", "28.02", "HO 16"
            Save_Form
        '-------
    '======
 
    'F7
    '======
        '-------
            Dim Dict As Dictionary
            Set Dict = Set_Dict("XS", "1900")
            Add_Size_QTY Dict
            Save_Form
        '-------
    '======
    
    'F8
    '======
        '-------
            Schedule "2016-07-18", "2016-07-26"
            Save_Form
        '-------
    '======
        
End Sub
'2016-07-26
Public Function Schedule(ReqDate As String, _
                         PromisedDate As String, _
                         Optional myrow As Integer, _
                         Optional Col As Dictionary)
'Dim ReqDate As String
'Dim PromisedDate As String
    
    'ReqDate = "2016-07-18"
    'PromisedDate = "2016-07-26"
    
    If GetESC Then SetESC
    Goto_Form "F8"
    SendMyData "%s"
    Schedule_Warning '0803
    Delay 300
    SendMyData ReqDate
    '-------
        If ReqDate <> "" And myrow <> 0 Then
            Cells(myrow, Col("Customer Req Date")).Interior.Color = vbGreen
        End If
    '-------
    SendMyData "Tab"
    SendMyData PromisedDate
    '-------
        If PromisedDate <> "" And myrow <> 0 Then
            Cells(myrow, Col("Promised Date")).Interior.Color = vbGreen
        End If
    '-------
    
    SendMyData "%o"
    
End Function

Private Function Schedule_Warning()

Dim hwnd As Long
Dim i As Integer
    Init
    hwnd = FindWindow(vbNullString, "Warning")
    i = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Exit Function
        End If
        
        If i = 6 Then
            Exit Do
        End If
        hwnd = FindWindow(vbNullString, "Warning")
        Delay 50
        i = i + 1
    Loop
    
    If hwnd <> 0 Then
        SendMyData "%y"
    End If
    
End Function
Private Sub test_Add_Size_QTY()
    'F7
    '======
        '-------
            Dim Dict As Dictionary
            'Set Dict = Set_Dict("XL REGULAR|L|XXL|S|XL|N", "150|400|400|300|250|400")
            Set Dict = Set_Dict("L PLUS", "1900")
            Add_Size_QTY Dict
            'Save_Form
        '-------
    '======
End Sub

Public Function Add_Size_QTY(Dict As Dictionary, _
                             Optional MP As Dictionary)
Dim i As Integer
Dim j As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Myhwnd As Long
Dim Size As String, QTY As String
    
    Call Init
    hwnd = setHwnd
    
    Goto_Form "F7"
    
    '------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------

Dim MyRect As RECT
Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer, x As Integer, y As Integer
Dim Flag As Integer
Dim hDC As Long
Dim MyColorNum As Long

            hDC = SetHdc
            
            '----��������ֵ���ж��Ƿ�����Size����
            Myhwnd = CLng(AllHwnd(128))
            Call GetWindowRect(Myhwnd, MyRect) '��ȡItem ID��ӦRectangle����
            x0 = MyRect.Left + 27
            y0 = MyRect.Top + 39
            Flag = 0
            For x = x0 To x0 + 10
                MyColorNum = GetPixel(hDC, x, y0)
                'Debug.Print MyColorNum
                If MyColorNum = 0 Then '0��ʾΪ��ɫ���������ݡ�
                    Flag = 1
                    Exit For
                End If
            Next
            '----
        
        If Flag = 0 Then '���Flag=1���򲻽������²�����
        
        '==============
            
            For i = 0 To Dict.Count - 1
            
                Init
                
                Size = CStr(Dict.Keys(i))
                QTY = CStr(Dict.Items(i))
                                
                
                If InStr(VBA.LCase(Size), "none") Then
                    SendMyData "%a"
                    SendMyData Size 'Size
                Else
                'ѡ��size����
                '===========
                    SendMyData "%n" '���New��ť
                    Item_Variations Size '���Size
                    Delay 300
                    
                    SendMyData "%a" '���Add
                    Sleep 300 '����ʱ,��ֹ���� 2016-09-27
                    '-----
                        If VBA.Len(Size) = 1 Then
                            SendMyData Size 'Size
                        Else
                            Do
                                DoEvents
                                Init
                                Sleep 300
                                For j = 1 To VBA.Len(Size)
                                    temp = VBA.Mid(Size, j, 1)
                                    SendMyData temp 'Size
                                    Myhwnd = CLng(AllHwnd(129)) '��ǰSizeλ��
                                    temp = GetText(Myhwnd) '��ȡϵͳsizeλ��ʾ���ı����ݡ�
                                    If VBA.LCase(temp) = VBA.LCase(Size) Then '�ж��Ƿ��Ҫ���Sizeƥ�䣬����ǣ�������ѭ��
                                        Exit Do
                                    End If
                                Next j
                            Loop
                        End If
                    '-----
                '===========
                End If
                
                If Size <> "" And Not MP Is Nothing Then
                    Cells(CInt(MP.Keys(i)), CInt(MP.Items(i))).Interior.Color = vbGreen
                End If
                    
                Delay 200

                SendMyData "Tab"
                SendMyData QTY 'QTY
                
                If QTY <> "" And Not MP Is Nothing Then
                    Cells(CInt(MP.Keys(i)), CInt(MP.Items(i)) + 1).Interior.Color = vbGreen
                End If
                
                Delay 300
            Next i
        '==============
        End If
    
End Function

Private Function Item_Variations(Size As String)

Dim i As Integer
Dim j As Integer
Dim t As Integer

Dim hwnd As Long
Dim Myhwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Data() As String
Dim MyRect As RECT
Dim Title As String
    
        If GetESC Then SetESC
        
        Title = "Item Variations"
        hwnd = FindWindow(vbNullString, Title)
        t = 0
        
        Delay 300
        
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
 
            
            Delay 500
            
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
            
            'Myhwnd = CLng(AllHwnd(17))
            'Click_By_Hwnd Myhwnd
            
            'Sizes
            '-------
                'Dim Size As String
                'Size = "XS"
                
                If UBound(AllHwnd) < 4 Then
                    ws_PopUp "Item Variations���ڲ�����ȡ����!"
                    End
                End If
                
                Myhwnd = CLng(AllHwnd(4))
                SetFocusAPI Myhwnd
                SendMyData "Right"
                'Add
                SendMyData "%a"
                Delay 200
                'Input size
                SendMyData Size
                '-------'Save
                    Myhwnd = CLng(AllHwnd(17))
                    Click_By_Hwnd Myhwnd 'Click Save
                    Sleep 200
                '-------
                If Invalid_Entry = True Then
                    SendMyData "Enter"
                    Delay 200
                    SendMyData "%d" 'Delete
                    '---
                        Delete_Confirmation ' Confirmation-->Yes
                    '---
                    '-------'Save
                        Myhwnd = CLng(AllHwnd(17))
                        Click_By_Hwnd Myhwnd 'Click Save
                        Sleep 200
                    '-------
                    SendMyData "%c" 'Close
                Else
                    Myhwnd = CLng(AllHwnd(4))
                    SetFocusAPI Myhwnd
                    SendMyData "Left"
                    Myhwnd = CLng(AllHwnd(6)) 'Create Defaults
                    Click_By_Hwnd Myhwnd
                    '-------'Save
                        Myhwnd = CLng(AllHwnd(17))
                        Click_By_Hwnd Myhwnd 'Click Save
                        Sleep 200
                    '-------
                    SendMyData "%c" 'Close
                End If
                
                Delay 200
            '-------
    
End Function
Private Function Invalid_Entry() As Boolean

Dim hwnd As Long
Dim i As Integer
Dim Title As String
    
    Invalid_Entry = False
    Title = "Invalid Entry"
    hwnd = FindWindow(vbNullString, Title)
    i = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Exit Function
        End If
        
        If i = 10 Then
            Exit Do
        End If
        hwnd = FindWindow(vbNullString, Title)
        Delay 50
        i = i + 1
    Loop
    
    If hwnd <> 0 Then
        Invalid_Entry = True
        Set ws = CreateObject("WSCRIPT.SHELL")
        ws.AppActivate Title
        Delay 200
    End If
    
End Function

Private Function Delete_Confirmation()

Dim hwnd As Long
Dim i As Integer
Dim Title As String
    
    Title = "Confirmation"
    hwnd = FindWindow(vbNullString, Title)
    i = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Exit Function
        End If
        
        If i = 10 Then
            Exit Do
        End If
        hwnd = FindWindow(vbNullString, Title)
        Delay 50
        i = i + 1
    Loop
    
    If hwnd <> 0 Then
        SendMyData "%y"
    End If
    
End Function
Private Sub test_Order_Lines()
        Order_Lines "Buy/Sell", "GAP-GKLBL-198585-G-CN", "1900", "28.02", "HO 16"
        Save_Form
End Sub
Public Function Order_Lines(OrderType As String, _
                            ItemNO As String, _
                            QTY As String, _
                            Price As String, _
                            Optional StyleNO As String, _
                            Optional myrow As Integer, _
                            Optional Col As Dictionary)
Dim i As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Myhwnd As Long
Dim HwndNum As Integer
Dim OffsetNum As Integer
Dim Data() As String
Dim StrData As String
    
    
    Call Init
    hwnd = setHwnd
    
    Goto_Form "F5"
    SendMyData "%a" 'Add
    Delay 500
    
    '------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------
    Debug.Print UBound(AllHwnd)
    
    
    HwndNum = UBound(AllHwnd)
    If HwndNum = 172 Then
        OffsetNum = 0
    Else
        OffsetNum = -1
    End If

 
    'Line Num
        Myhwnd = CLng(AllHwnd(68 + OffsetNum))
        temp = GetText(Myhwnd)
        If temp <> "" Then
            Data = VBA.Split(VBA.LCase(temp), "of")
            If UBound(Data) > 0 Then
                temp = VBA.Trim(Data(1))
                temp = VBA.Replace(temp, VBA.Chr(0), "")
            End If
        End If
        
    'Order Type
        'Dim OrderType As String
        'OrderType = "Buy/Sell"
        If OffsetNum = 0 Then
            Myhwnd = CLng(AllHwnd(84)) '���Գ�Order Type�ľ��ֵ
            SendStrToSys Myhwnd, OrderType '��������
            Delay 500
            Do While (Invalid_Selection) '�ж��Ƿ��жԻ��򵯳�
                SendMyData "Enter"
                Delay 800
                Myhwnd = CLng(AllHwnd(84))
                SendStrToSys Myhwnd, OrderType
                Delay 800
            Loop
        End If
        
        
    'RVL Item No
        'Dim ItemNo As String
        'ItemNo = "GAP-GKLBL-198585-G-CN"
        If ItemNO <> "" Then
            Myhwnd = CLng(AllHwnd(82 + OffsetNum)) 'Delete
            SetFocusAPI Myhwnd
            For i = 1 To 4
                SendMyData "Tab"
            Next i
            Delay 200
            SendMyData ItemNO
            SendMyData "Enter"
            Call Item_Search '�������Item Search�Ի���Ĭ�ϵ�һ��
            Delay 300
        End If
        
        If ItemNO <> "" And myrow <> 0 Then
            Cells(myrow, Col("Item NO.")).Interior.Color = vbGreen
        End If
        
    'Order Qty
        'Dim QTY As String
        'QTY = "1900"
        If QTY <> "" Then
            Myhwnd = CLng(AllHwnd(82 + OffsetNum)) 'Delete
            SetFocusAPI Myhwnd
            
            For i = 1 To 6
                SendMyData "Tab"
            Next i
            Delay 300
            SendMyData QTY
            SendMyData "Tab"
            Delay 200
        End If
        If QTY <> "" And myrow <> 0 Then
            Cells(myrow, Col("Total")).Interior.Color = vbGreen
        End If
        
    'Unit Price
        'Dim Price As String
        'Price = "28.02"
        If Price <> "" Then
            SendMyData Price
            SendMyData "Tab"
            Delay 200
        End If
        
        If Price <> "" And myrow <> 0 Then
            Cells(myrow, Col("Unit Price")).Interior.Color = vbGreen
        End If
        
    'Style NO.
        'Dim StyleNO As String
        'StyleNO = "HO 16"
        If StyleNO <> "" Then
            For i = 1 To 3
                SendMyData "Tab"
            Next i
            Delay 300
            SendMyData StyleNO
        End If
        
        If StyleNO <> "" And myrow <> 0 Then
            Cells(myrow, Col("Style NO.")).Interior.Color = vbGreen
        End If
        
End Function
Private Function Item_Search() '(ItemNO As String)
'Dim ItemNO As String
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
    
    Delay 800
    
    Title = "Item Search"
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
        If t > 10 Then 'ѭ��1.5��
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
    If UBound(AllHwnd) > 0 Then
        'Select
        '----
            Myhwnd = CLng(AllHwnd(4)) 'Select
            PostMessage Myhwnd, BM_CLICK, 0, 0 '���Select��ť.
        '----
        Delay 300
    End If
    '--------------

End Function
Private Function Invalid_Selection() As Boolean

Dim hwnd As Long
Dim i As Integer
Dim Title As String

    If GetESC Then SetESC
    
    Invalid_Selection = False
    Title = "Invalid Selection"
    hwnd = FindWindow(vbNullString, Title)
    i = 0
    Do While hwnd = 0
        DoEvents
        If GetESC Then
            Exit Function
        End If
        
        If i = 10 Then
            Exit Do
        End If
        hwnd = FindWindow(vbNullString, Title)
        Delay 50
        i = i + 1
    Loop
    
    If hwnd <> 0 Then
        Invalid_Selection = True
    End If
    
End Function
Private Function SendStrToSys(hwnd As Long, MyData As String) '0725
    
    Dim i As Integer
    Dim StrData As String
        
        For i = 1 To VBA.Len(MyData)
            StrData = VBA.Mid(MyData, i, 1)
            PostMessage hwnd, WM_CHAR, VBA.Asc(StrData), 0
        Next i
        Delay 200
        
End Function


Public Function Ship_To_Address(MyData As String)
'Dim MyData As String
    
    'MyData = "4596"
    If GetESC Then SetESC
    Call Init
    Goto_Form "F4"
    SendMyData "%a"
    SendMyData MyData
    SendMyData "Enter"
    
End Function
Private Function test_Bill_to_Customer()
    Bill_to_Customer 6200
End Function
Public Function Bill_to_Customer(MyData As String)
Dim i As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Myhwnd As Long

    
    Call Init
    hwnd = setHwnd
    
    '------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------

        Myhwnd = CLng(AllHwnd(27)) '���Կ�֪ΪTabControl�ؼ�
        
        SetFocusAPI Myhwnd
        
        For i = 1 To 2
            SendMyData "Tab"
        Next
        
        SendMyData MyData
        
        SendMyData "Enter"
    
End Function

Public Function Customer_PO_NO(MyData As String)
Dim i As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Myhwnd As Long

    
    Call Init
    hwnd = setHwnd
    
    '------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------

        Myhwnd = CLng(AllHwnd(27)) 'TabControl�ؼ�
        
        SetFocusAPI Myhwnd '��ȡ�������
        
        For i = 1 To 5
            SendMyData "Tab"
        Next
        
        SendMyData MyData
            
End Function
Public Function Order_Recd_Date(MyData As String)
Dim i As Integer
Dim hwnd As Long
Dim AllHwnd() As String
Dim temp As String
Dim Myhwnd As Long

    
    Call Init
    hwnd = setHwnd
    
    '------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        AllHwnd = VBA.Split(ChildHwnd, ",")
    '----------------------

        Myhwnd = CLng(AllHwnd(27)) 'TabControl�ؼ�
        
        SetFocusAPI Myhwnd '��ȡ�������
        
        For i = 1 To 9
            SendMyData "Tab"
        Next
        
        SendMyData MyData
End Function
Public Function Save_Form()

    Call Init
    SendMyData "^s"
    Delay 1000
    
End Function
 
Private Function Goto_Form(MyData As String)
 
    Call Init
    
    Select Case VBA.UCase(MyData)
        Case "F3"
            SendMyData "{F3}"
        Case "F4"
            SendMyData "{F4}"
        Case "F5"
            SendMyData "{F5}"
        Case "F6"
            SendMyData "{F6}"
        Case "F7"
            SendMyData "{F7}"
        Case "F8"
            SendMyData "{F8}"
        Case "F9"
            SendMyData "{F9}"
    End Select
    
    Delay 1000

End Function

Private Function Click_By_Hwnd(Myhwnd As Long)

    PostMessage Myhwnd, BM_CLICK, 0, 0 '���Select��ť.
        
End Function

'2016-07-28
Public Function Popup_Search(AddressNO As String)
'Dim AddressNO_temp As String
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
    
    If UBound(AllHwnd) <= 0 Then
        Exit Function
    End If

Dim x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer
Dim Path As String
Dim Flag As Integer
Dim PicStr As String
Dim myrow As Integer
Dim TargetRow As Integer
Dim Times As Integer
Dim DownFlag As Integer
Dim MyLastRow As Integer
 

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
        MyLastRow = Get_Last_Row(x0, y0, x1, y1)
 
        
        Dim MyRange As ScreenRange
        Dim Tdata() As String

'        '-------
            'Dim AddressNO As String
            'AddressNO = "1655/6" '"4596/11" ' "6200/1" ' "1434/12" '4596/9 6200/1
'        '--------

        'stime = Timer
        
        Erase Data
        Erase Tdata
        i = 1
        Flag = 0
        Do
            DoEvents
            If GetESC Then
                Debug.Print "All End"
                End
            End If
            
            MyRange = GetRange(i, x0, y0) '��ȡ��ɫ��
            PicStr = Get_Picture_Str(Path, MyRange.Left, MyRange.Top, MyRange.Width, MyRange.Height, 2, 2)
            temp = PicStr
            temp = VBA.Replace(temp, " ", "")
            temp = VBA.Replace(temp, VBA.Chr(13), "")
            temp = VBA.Replace(temp, VBA.Chr(10), "")
            If AddressNO = temp Then
                TargetRow = (i - 1)
                Debug.Print temp
                Flag = 1
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
        
        If Flag = 0 Then
            MsgBox "û���ҵ����ݣ�" & AddressNO
            End
        End If

        'Select
        '----
            Myhwnd = CLng(AllHwnd(2)) 'Select
            PostMessage Myhwnd, BM_CLICK, 0, 0 '���Select��ť.
        '----
 
    '--------------
End Function

Private Function Select_Row(TargetRow As Integer, myrow As Integer)
Dim Times As Integer
Dim j As Integer
Dim t As Integer

    'ѡ�ж�Ӧ����
    '-----
        Times = myrow - 1
        t = TargetRow - Times
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
End Function
Private Function Get_Last_Row(x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer) As Integer
Dim hDC As Long
Dim MyColorNum As Long
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer
Dim temp As String
Dim Data
Dim myrow As Integer

    hDC = SetHdc
    
    Data = Array(10, 30, 48, 67, 86, 106, 125) 'ÿ����ֵ
    
    For i = 0 To UBound(Data)
    
        y = y0 + Data(i)
        
        DoEvents

        MyColorNum = GetPixel(hDC, x0, y)
        If MyColorNum = 16750899 Then '��ɫ��
            myrow = GetMyRow(y0, y) '�ڵڼ���
            If InStr(temp, myrow) = 0 Then
                For x = x0 + 10 To x0 + 20
                    MyColorNum = GetPixel(hDC, x, y)
                    If MyColorNum = 16777215 Then ' MyColorNum = 16777215 ��ɫ������Ϊ��ɫ
                        temp = temp & myrow
                        Exit For
                    End If
                Next

            End If
        Else
            myrow = GetMyRow(y0, y) '�ڵڼ���
            If InStr(temp, myrow) = 0 Then
                For x = x0 + 10 To x0 + 20
                    MyColorNum = GetPixel(hDC, x, y)
                    If MyColorNum = 0 Then ' MyColorNum = 0 ����Ϊ��ɫ
                        temp = temp & myrow
                        Exit For
                    End If
                Next
            End If
        End If
        
        If VBA.Len(temp) > 1 Then
            If myrow - VBA.Len(temp) > 1 Then '�ж��Ƿ�հ׶��У�����ǣ�������ѭ��������ʱ�䣬����ÿ�ζ��ж�7������
                Exit For
            End If
        End If
        
    Next
    
    Get_Last_Row = CInt(VBA.Right(temp, 1))
    
End Function
Private Function Get_Blue_Row(x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer) As Integer
Dim hDC As Long
Dim MyColorNum As Long
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer

    hDC = SetHdc
    'Debug.Print y0 160
    'Debug.Print y1 313
    For y = y0 To y1 Step 9
        DoEvents
        'Call SetCursorPos(x0, y)
        MyColorNum = GetPixel(hDC, x0, y)
        'Debug.Print MyColorNum
        If MyColorNum = 16750899 Then '��ɫ��16750899 ��ɫ��16777215
            Get_Blue_Row = GetMyRow(y0, y)
            Exit For
        End If
        'Delay 20
    Next
    
End Function

Private Function test_GetMyRow()
    Debug.Print GetMyRow(181)
End Function
Private Function GetMyRow(y0 As Integer, Num As Integer) As Integer
'��Ļȡֵ

    Select Case Num 'y0 = 160
        Case (y0 + 1) To (y0 + 19) '161 To 179 10
            GetMyRow = 1
            
        Case (y0 + 21) To (y0 + 39) '181 To 198 30
            GetMyRow = 2
            
        Case (y0 + 40) To (y0 + 57) '200 To 217 49
            GetMyRow = 3
            
        Case (y0 + 59) To (y0 + 76) '219 To 236 68
            GetMyRow = 4
            
        Case (y0 + 78) To (y0 + 95) '238 To 255 87
            GetMyRow = 5
            
        Case (y0 + 97) To (y0 + 114) '257 To 274 106
            GetMyRow = 6
            
        Case (y0 + 116) To (y0 + 134) '276 To 294 125
            GetMyRow = 7
            
        Case (y0 + 20), (y0 + 40), (y0 + 58), (y0 + 77), (y0 + 96), (y0 + 115)
            MsgBox "��ĻȡֵΪ�߽�ֵ,�޷��ж�,����ֹͣ��"
            End
        Case Else
            MsgBox "������Ļȡֵ��Χ,����ֹͣ��"
            End
    End Select
    
End Function

Private Function SetHdc()
Dim hDC As Long
Dim dm As DEVMODE

    hDC = CreateDC("DISPLAY", "", "", dm)
    SetHdc = hDC

End Function

Public Function Get_Picture_Str(Path As String, Left, Top, Width, Height, Optional MultipleW = 1, Optional MultipleH = 1, Optional dwRopFlag As Integer = 1)
              
    Dim temp As String
    Dim Times As Integer
        
        Times = 0
        Do
            DoEvents
            
            '��Ļ��ͼ
            GetPrintScreen Path, Left, Top, Width, Height, MultipleW, MultipleH, dwRopFlag '18  10
            
            'OneNote
            'ͼƬ������ȡ
            temp = Image_Str(Path, OneNote)
            
            
            If temp <> "Image is not readable" Then
                Get_Picture_Str = temp
                Exit Do
            End If
            
            If Times > 2 Then
                Get_Picture_Str = temp
                Exit Do
            End If
            Times = Times + 1
            Delay 100
        Loop
        
End Function


Public Function Master_Customer(MyData As String)
Dim hwnd As Long
Dim AllHwnd() As String

Dim i As Integer
Dim t As Long
Dim EditHwnd As Long
Dim temp As String
Dim StrData As String
Dim Myhwnd As Long

    
        Call Init '����ϵͳ����
        hwnd = setHwnd '���
        
        '-------------------
        '��ȡ�����ڵ��ӿؼ����
            ChildHwnd = ""
            Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&) 'ö���Ӵ���
            ChildHwnd = VBA.Mid(ChildHwnd, 2)
            AllHwnd = VBA.Split(ChildHwnd, ",")
        '---------------------

        '0722 Master Customer
        '-----------
        '��ʼ��Sales Order View ��Index��New �򿪵�Sales Order View��Index�ǲ�һ����
            Dim MasterHwndIndex As Integer
            Myhwnd = CLng(AllHwnd(49))
            temp = FGetClassName(Myhwnd)
            If VBA.LCase(temp) = "edit" Then
                MasterHwndIndex = 49
            Else
                MasterHwndIndex = 44
            End If
            '---------����Master Customer
                Myhwnd = CLng(AllHwnd(MasterHwndIndex))
                temp = MyData ' "1655"
                For i = 1 To Len(temp)
                    StrData = VBA.Mid(temp, i, 1)
                    PostMessage Myhwnd, WM_CHAR, Asc(StrData), 0
                Next i
                Delay 100
                PostMessage Myhwnd, WM_KEYDOWN, VK_ENTER, 0
                PostMessage Myhwnd, WM_KEYUP, VK_ENTER, 0
                Delay 200
            '---------
       '-----------
       
       '---


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
Dim FormStr As String

        Path = "C:\1\1.bmp"
        If Dir("C:\1\", vbDirectory) = "" Then
            MkDir "C:\1"
        End If
    '--------------
        Myhwnd = CLng(AllHwnd(27))
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '��ȡ��ӦRectangle����,Order No.�ľ�������
        '===================
            
            FormStr = "F3"
            
            Select Case FormStr
                Case "F3"
                    x0 = MyRect.Left + 179 '132
                    y0 = MyRect.Top + 42 '41
                    PicStr = Get_Picture_Str(Path, x0, y0, 46, 14, 2, 2)   '����94,��15��W�Ŵ�����4,H�Ŵ�����6
                Case "F4", "F5"
                    x0 = MyRect.Left + 356 '
                    y0 = MyRect.Top + 38 '
                    PicStr = Get_Picture_Str(Path, x0, y0, 46, 14, 2, 2)
            End Select
 
        '===================
        temp = PicStr
        temp = VBA.Replace(temp, VBA.Chr(0), "")
        If temp = "Image is not readable" Or VBA.Trim(temp) = "" Then
            PicStr = ""
        End If
        
        PicStr = VBA.Replace(PicStr, VBA.Chr(0), "")
        PicStr = VBA.Replace(PicStr, " ", "")
        PicStr = VBA.Replace(PicStr, VBA.Chr(10), "")
        PicStr = VBA.Replace(PicStr, VBA.Chr(13), "")
        Get_OrderNO = PicStr
        
        'Debug.Print PicStr
    '--------------
End Function







 
