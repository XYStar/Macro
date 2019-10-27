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


Dim XY As POINTAPI
Dim AppData As String

'===========2016============


'����һ���ص�����, �������ģ����. ��������ָ�����ڵ��Ӵ���(�ؼ�). ��������е� hWnd ��Ϊ�Ӵ���(�ؼ�)���
Private Function EnumChildProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
ChildHwnd = ChildHwnd & "," & hwnd
EnumChildProc = 1
End Function


' ����: FGetClassName
' ����: ����ָ�������е�����
' ����: hWnd ָ�����ڵľ��
' ����: ָ�����ڵ�����

Private Function FGetClassName(hwnd As Long) As String
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
        'ȥ���������Ч�ַ�
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        'ת��������
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
 

        'Ctrl+Wheel �Ŵ�
        '==============
            DoEvents
            Ctrl_Wheel 1
            DoEvents
            Sleep 200
        '===============
    
    Flag = 0
    Do
        DoEvents
        Call GetWindowRect(Myhwnd, MyRect) '��ȡ��ӦRectangle����,Order No.�ľ�������
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
            Flag = Flag + 1 '�����ȡ������Ϊ��,�����ȡһ��
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
        
        If Flag = 2 Then '������ν����Ϊ�գ�������ѭ��
            Exit Do
        End If
    Loop
    '--------------
    
    'Ctrl+Wheel ��С,����ԭ����С
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
        'ѡItem
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
        Sleep 300
        '------------
    '----------------------------------------
        ChildHwnd = ""
        Call EnumChildWindows(hwnd, AddressOf EnumChildProc, ByVal 0&)
        'ȥ���������Ч�ַ�
        ChildHwnd = VBA.Mid(ChildHwnd, 2)
        'ת��������
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
        
        SetFocusAPI Myhwnd '��ȡ����
        
        'Ctrl+Wheel �Ŵ�
        '==============
            DoEvents
            Ctrl_Wheel 4
            DoEvents
            Sleep 100
        '===============


        Call GetWindowRect(Myhwnd, MyRect) '��ȡ��ӦRectangle����,Order No.�ľ�������
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
        'ѡ��Item ID �㷨
            Do
                
                ws.AppActivate Title
                
                DoEvents
                
                'ѡ�е�һ��
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
                    MsgBox "����ִ�в��ɹ�,�����³���!", , "Image is not readable"
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
                    MyColorNum = GetPixel(hDC, x1, y1) '�ж����½��Ƿ��й���������
                    DeleteDC hDC
                    '--------
                    
                    If MyColorNum <> 16777215 Then '��ʾ�й��������֣�����Ҫ��PageDown��ҳ
                        SendMyData "PgDn"
                        Sleep 200
                    Else
                        Debug.Print "û���ҵ�Item:" & Item
                        MsgBox "û���ҵ�Item:" & Item
                        End
                    End If
                End If
            Loop
        '###########################
        
        
        'Select
        '----
            SendMyData "%s"
            PO_Warning '�ж��Ƿ��жԻ��򵯳�
            Sleep 500
        '----
        
End Function



Private Function My_Set() As Boolean
    '���ڲ���
    '============
    '2016-01-29
        Dim sys_infm, UserName, Pcname
        Set sys_infm = CreateObject("WSCRIPT.NETWORK")
        UserName = sys_infm.UserName
        Pcname = sys_infm.computername
        
        My_Set = False
                
        '���������ʱ������ʱ����
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
            ws_PopUp "û���ҵ�:Comments����!"
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
            MsgBox "û���ҵ�Comment��!"
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

        Myhwnd = CLng(AllHwnd(24)) '���Կ�֪��AllHwnd(24)ΪTabControl�ؼ�
        
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
        MsgBox "û���ҵ�:" & Title
    Else
        Set ws = CreateObject("WSCRIPT.SHELL")
        ws.AppActivate Title
    End If
    
    mouse_event MOUSEEVENTF_LEFTDOWN, x, y, 0, 0
    Sleep 20
    mouse_event MOUSEEVENTF_LEFTUP, x, y, 0, 0
    
End Function
  
