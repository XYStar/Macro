Attribute VB_Name = "All_Set_Cursor"
'�������״̬
'-----------------
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetSystemCursor Lib "user32" (ByVal hcur As Long, ByVal id As Long) As Long
Private Declare Function LoadCursorFromFile Lib "user32" Alias "LoadCursorFromFileA" (ByVal lpFilename As String) As Long

Private Const OCR_NORMAL = 32512

'�ж�Windowsϵͳ�汾
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONONFO) As Long

Private Type OSVERSIONONFO
dwOSVersioninfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformld As Long
dwCSDVersion As String * 128
End Type
'
'----------------------------------


Private Sub test_SetMyCursor()
    Call SetMyCursor(0)
End Sub

''#################################################################################################################################################
 
'�ı�ϵͳ���״̬
'
''#################################################################################################################################################

'�ı����ָ����״
Public Function SetMyCursor(i As Integer) '2014-08-21 by Star
Dim sPath As String * 260   '����·��
Dim PathData As String
Dim Path As String
Dim CursorData(4) As String
Dim hCursor As Long
Dim StrData As String

'���ָ��
'---------------------------------------------------------
CursorData(0) = "\aero_arrow.cur" '����. C:\windows\Cursors\aero_arrow.ani
CursorData(1) = "\aero_busy.ani" 'Wait C:\windows\Cursors\aero_busy.ani
CursorData(2) = "\arrow_m.cur" ''  "\3dwarro.cur" '����. C:\windows\Cursors\arrow_m.cur
CursorData(3) = "\busy_m.cur" 'Wait C:\windows\Cursors\busy_m.cur
GetSystemDirectory sPath, Len(sPath)
PathData = Replace(sPath, VBA.Chr(0), "")

Path = Replace(PathData, "system32", "Cursors") ' C:\windows\Cursors\

StrData = SystemVer
'---------------------------------------------------------
If i = 0 Then
    If StrData = "Windows 7" Then
        '���ָ��ָ�����..
        '---------------------------------------
        PathData = Path & CursorData(0) ''����. C:\windows\Cursors\aero_arrow.ani
        hCursor = LoadCursorFromFile(PathData)
        Call SetSystemCursor(hCursor, OCR_NORMAL)
        '---------------------------------------
    Else
        '���ָ��ָ�����..
        '---------------------------------------
        PathData = Path & CursorData(2) ''����. C:\windows\Cursors\arrow_m.ani
        hCursor = LoadCursorFromFile(PathData)
        Call SetSystemCursor(hCursor, OCR_NORMAL)
        '---------------------------------------
    End If
End If
If i = 1 Then
    If StrData = "Windows 7" Then
        '���ָ�뿪ʼ�ȴ�..
        '---------------------------------------
        PathData = Path & CursorData(1) ''����. C:\windows\Cursors\aero_arrow.ani
        hCursor = LoadCursorFromFile(PathData)
        Call SetSystemCursor(hCursor, OCR_NORMAL)
        '---------------------------------------
    Else
        '���ָ�뿪ʼ�ȴ�..
        '---------------------------------------
        PathData = Path & CursorData(3) ''����. C:\windows\Cursors\busy_m.ani
        hCursor = LoadCursorFromFile(PathData)
        Call SetSystemCursor(hCursor, OCR_NORMAL)
        '---------------------------------------
    End If
End If

    DoEvents

End Function

''#################################################################################################################################################
 
'��ȡ����ϵͳ�汾��
'
''#################################################################################################################################################

Private Function SystemVer() As String '��ȡ����ϵͳ�汾�� by star 2014-09-18

Dim Osinfor As OSVERSIONONFO
Dim StrOsName As String
Osinfor.dwOSVersioninfoSize = Len(Osinfor)
GetVersionEx Osinfor
Select Case Osinfor.dwPlatformld
       Case 0
            StrOsName = "Windows 32s"
       Case 1
          Select Case Osinfor.dwMinorVersion
                 Case 0
                      StrOsName = "Windows 95"
                 Case 10
                      StrOsName = "Windows 98"
                 Case 90
                      StrOsName = "Windows Mellinnium"
          End Select
       Case 2
          Select Case Osinfor.dwMajorVersion
                 Case 3
                      StrOsName = "WindowsNT 3.51"
                 Case 4
                      StrOsName = "WindowsNT 4.0"
                 Case 5
                      Select Case Osinfor.dwMinorVersion
                             Case 0
                                  StrOsName = "Windows 2000"
                             Case 1
                                  StrOsName = "Windows XP"
                             Case 2
                                  StrOsName = "Windows 2003"
                      End Select
                 Case 6
                      Select Case Osinfor.dwMinorVersion
                             Case 0
                                  StrOsName = "Windows Vista"
                             Case 1
                                  StrOsName = "Windows 7"
                      End Select
         End Select
       Case Else
            StrOsName = "δ֪ϵͳ�汾"
       End Select
       SystemVer = StrOsName
End Function
