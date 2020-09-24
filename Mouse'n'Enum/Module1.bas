Attribute VB_Name = "Module1"
' *************************************
' * Geschrieben von                   *
' *     Wolfgang Ehrhardt             *
' *         woeh@gmx.de               *
' *                                   *
' * Der Code ist frei für Jedermann,  *
' * solange der Quellcode NICHT für   *
' * kommerzielle Zwecke verwendet     *
' * wird.                             *
' *                                   *
' * Sollten Teile oder Auszüge aus    *
' * diesem Quellcode für kommerzielle *
' * Zwecke verwendet werden, bitte    *
' * ich um Kontaktaufnahme unter      *
' *         woeh@gmx.de               *
' *                                   *
' * Vielen Dank an ActiveVB           *
' * & das ActiveVB-Forum              *
' *************************************
Declare Function SetParent Lib "USER32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function GetParent Lib "USER32" (ByVal Hwnd As Long) As Long
Declare Function GetWindow Lib "USER32" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
Declare Function IsWindowVisible Lib "USER32" (ByVal Hwnd As Long) As Long
Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetWindowThreadProcessId Lib "USER32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindowRect Lib "USER32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Declare Function SearchLB Lib "USER32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function GetMenu Lib "USER32" (ByVal Hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "USER32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemInfo Lib "USER32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Declare Function GetMenuItemID Lib "USER32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Declare Function GetSystemMenu Lib "USER32" (ByVal Hwnd As Long, ByVal bRevert As Long) As Long
Declare Function GetSubMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetWindowLongA Lib "user32.dll" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Declare Function EnableWindow Lib "USER32" (ByVal Hwnd As Long, ByVal fEnable As Long) As Long
Declare Function SetWindowPos Lib "USER32" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32.dll" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SetForegroundWindow Lib "USER32" (ByVal Hwnd As Long) As Long
Declare Function GetDC Lib "USER32" (ByVal Hwnd As Long) As Long
Declare Function RemoveMenu Lib "USER32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function PostMessage Lib "USER32" Alias "PostMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function MoveWindow Lib "USER32" (ByVal Hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function BringWindowToTop Lib "USER32" (ByVal Hwnd As Long) As Long
Declare Function SendMessageByNum& Lib "USER32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function WindowFromPoint Lib "USER32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Declare Sub mouse_event Lib "USER32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function EnumWindows Lib "USER32" (ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function MoveForm Lib "USER32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "USER32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Private Declare Function EnumChildWindows Lib "USER32" (ByVal hwndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Any) As Long
Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
Private Declare Function GetClassName& Lib "USER32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Private Declare Function SendMessageByString Lib "USER32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal CB As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal CB As Long, ByRef cbNeeded As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function CallWindowProc Lib "USER32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FlashWindowEx Lib "USER32" (pfwi As FLASHWINFO) As Boolean

Private Type FLASHWINFO
    cbSize As Long
    Hwnd As Long
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type

Public Type MainMenu1
    SaveMenu  As Integer
    SavePos   As Integer
    ToolTip   As Integer
    ColorP    As Integer
    StartMin  As Integer
    AllMenu   As Integer
End Type
    
Public Type MyPrg
    Name     As String
    INI      As String
    EXE      As String
    Path     As String
    Mail     As String
    ActiveVB As String
    VBtutor  As String
    OS       As Long
    OSstring As String
    TaskID   As Long
End Type

Public Type MycChilds
    Show    As Boolean
    Stopped As Boolean
End Type

Public Type MyMenuItem
    Type    As Long
    Text    As String
    Checked As Boolean
    Grayed  As Boolean
    Enabled As Boolean
    Hwnd    As Long
    ID      As Long
    MItem   As Long
    tHwnd   As Long
    AllType As Long
    SysMenu As Boolean
    Owner   As Long
End Type

Public Type MymMenu
    Show    As Boolean
    Stopped As Boolean
End Type

Public Type EnWin
    Hwnd        As Long
    Class       As String
    TaskID      As Long
    Text        As String
    Left        As Long
    Top         As Long
    Width       As Long
    Heigth      As Long
    ParentHwnd  As Long
    ParentClass As String
    ParentText  As String
    Visible     As Boolean
    Thread      As Long
    TopParent   As Long
End Type

Public Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wID           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String '* 255
    cch           As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Private Type PROCESSENTRY32
    dwSize              As Long
    cntUsage            As Long
    th32ProcessID       As Long
    th32DefaultHeapID   As Long
    th32ModuleID        As Long
    cntThreads          As Long
    th32ParentProcessID As Long
    pcPriClassBase      As Long
    dwFlags             As Long
    szExeFile           As String * 260
End Type

Public Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

Public Type MyIcon
    Hand    As StdPicture
    Move    As StdPicture
    Pointer As Integer
End Type

Public Const FLASHW_STOP = 0
Public Const FLASHW_CAPTION = &H1
Public Const FLASHW_TRAY = &H2
Public Const FLASHW_ALL = (FLASHW_CAPTION Or FLASHW_TRAY)
Public Const FLASHW_TIMER = &H4
Public Const FLASHW_TIMERNOFG = &HC

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Public Const GWL_EXSTYLE As Long = -20
Public Const GWL_STYLE = -16
Public Const GWL_WNDPROC = (-4&)

Public Const hNull = 0

Public Const HTCAPTION = 2

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const LB_ITEMFROMPOINT = &H1A9

Public Const MAX_PATH = 260

Public Const MF_BITMAP = &H4&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CHANGE = &H80&
Public Const MF_CHECKED = &H8&
Public Const MF_DEFAULT = &H1000&
Public Const MF_DELETE = &H200&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_POPUP = &H10&
Public Const MF_REMOVE = &H1000&
Public Const MF_RIGHTJUSTIFY = &H4000&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&

Public Const MFT_RADIOCHECK = &H200&
Public Const MFT_RIGHTORDER = &H2000&

Public Const MIIM_BITMAP = &H80
Public Const MIIM_CHECKMARKS = &H8
Public Const MIIM_DATA = &H20
Public Const MIIM_FTYPE = &H100
Public Const MIIM_ID = &H2
Public Const MIIM_STATE = &H1
Public Const MIIM_STRING = &H40
Public Const MIIM_SUBMENU = &H4
Public Const MIIM_TYPE = &H10

Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
Public Const MOUSEEVENTF_MIDDLEUP = &H40
Public Const MOUSEEVENTF_MOVE = &H1
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_RIGHTDOWN = &H8
Public Const MOUSEEVENTF_RIGHTUP = &H10

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1

Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_VM_READ = 16

Public Const RGN_XOR = 3

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_SHOWWINDOW = &H40

Public Const SYNCHRONIZE = &H100000

Public Const TH32CS_SNAPPROCESS = &H2&

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2

Public Const WM_ACTIVATE = &H6
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONUP = &H208
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEMOVE = &H200
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const WS_CHILD = &H40000000
Public Const WS_CAPTION = &HC00000
Public Const WS_OVERLAPPED = &H0
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_SYSMENU = &H80000
Public Const WS_HSCROLL = &H100000
Public Const WS_VSCROLL = &H200000
Public Const WS_DLGFRAME = &H400000
Public Const WS_BORDER = &H800000
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_DISABLED = &H8000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_MINIMIZE = &H20000000
Public Const WS_POPUP = &H80000000
Public Const WS_THICKFRAME = &H40000

Public Const WS_EX_LEFT = &H0
Public Const WS_EX_DLGMODALFRAME = &H1
Public Const WS_EX_NOPARENTNOTIFY = &H4
Public Const WS_EX_TOPMOST = &H8
Public Const WS_EX_ACCEPTFILES = &H10
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_WINDOWEDGE = &H100
Public Const WS_EX_CLIENTEDGE = &H200
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_RIGHT = &H1000
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_LEFTSCROLLBAR = &H4000
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000
Public Const WS_EX_APPWINDOW = &H40000

Private TickCount As Long

Public mY As MyPrg, Cchilds As MycChilds, MyMenu As MymMenu
Public eWin As EnWin, fIcon As MyIcon, MainMenu As MainMenu1

Public wHwnd As Long, Win2Find As Long, PrevWndProc As Long

Public View(6) As Boolean, Infos(4) As Boolean
Public Cmenu(3) As Boolean, CapView(9) As Boolean
Public tMenu(4) As Boolean
Public Function GetFromINI(AppName As String, KeyName As String, _
                           INI As String) As String
   Dim RetStr As String
   
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName, _
                                                     ByVal KeyName, "", _
                                                     RetStr, Len(RetStr), _
                                                     INI))
End Function
Public Sub Write2INI(Sektion As String, Abschnitt As String, _
                     Wert As String, INI As String)
    Call WritePrivateProfileString(Sektion, Abschnitt, Wert, INI)
End Sub
Public Sub ClickIcon(Icon As Long)
    Call SendMessage(Icon, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessage(Icon, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub FormMove(Frm As Form)
    Call ReleaseCapture
    Call MoveForm(Frm.Hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub
Private Sub GetPressedKey()
    Static Count As Integer
    Dim Index As Long
    Dim nPoint As POINTAPI
    
    If GetAsyncKeyState(VK_F12) Then Call GetCursorPos(nPoint): _
                                     Dummy.mX = nPoint.X: _
                                     Dummy.mY = nPoint.Y: _
                                     Call Dummy.PopUp(eWin.Hwnd)
    
    If GetAsyncKeyState(VK_F11) And Cchilds.Show Then
        Count = Count - 1
        If Count < 1 Then
            Index = FindForm("Ac", "cRefresh")

            If Cchilds.Stopped Then
                Call Forms(Index).AutoRefresh(True)
            Else
                Call Forms(Index).AutoRefresh(False)
            End If
            Count = 10
        End If
    End If
    
    If GetAsyncKeyState(VK_F10) And MyMenu.Show Then
        Count = Count - 1
        If Count < 1 Then
            Index = FindForm("Ac", "mRefresh")
            If MyMenu.Stopped Then
                Call Forms(Index).AutoRefresh(True)
            Else
                Call Forms(Index).AutoRefresh(False)
            End If
            Count = 10
        End If
    End If
    
End Sub
Sub TimerProc(ByVal Hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Call GetPressedKey
End Sub
Public Sub StayOnTop(Frm As Form)
    Call SetWindowPos(Frm.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                                           SWP_NOSIZE Or SWP_NOMOVE)
End Sub
Public Sub StayWinOnTop(Hwnd As Long, Remove As Boolean)
    Dim Command As Long
    
    Command = HWND_TOPMOST
    If Remove Then Command = HWND_NOTOPMOST
    
    Call SetWindowPos(Hwnd, Command, 0, 0, 0, 0, _
                                           SWP_NOSIZE Or SWP_NOMOVE)
    
End Sub
Public Function GetClass(Child As Long) As String
    Dim Buffer As String
    Dim GetClas As Long
    
    Buffer = Space(250)
    GetClas = GetClassName(Child, Buffer, 250)
    GetClass = Left(Buffer, GetClas)
    
End Function
Public Function GetText(Child As Long) As String
    Dim GetTrim As Long
    Dim TrimSpace As String
    
    GetTrim = SendMessageByNum(Child, WM_GETTEXTLENGTH, 0&, 0&)
    TrimSpace = Space$(GetTrim)
    GetString = SendMessageByString(Child, WM_GETTEXT, _
                                    GetTrim + 1, TrimSpace)
    GetText = TrimSpace
        
End Function
Public Function FindChildByClass(ParentW As Long, ChildHand As String)
    Dim Firs As Long, Firss As Long
    
    Firs = GetWindow(ParentW, GW_CHILD)
    If UCase(mID(GetClass(Firs), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo bone
    Firs = GetWindow(ParentW, GW_CHILD)
    If UCase(mID(GetClass(Firs), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo bone
    
    While Firs
        Firss = GetWindow(ParentW, GW_CHILD)
        If UCase(mID(GetClass(Firss), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo bone
        Firs = GetWindow(Firs, GW_HWNDNEXT)
        If UCase(mID(GetClass(Firs), 1, Len(ChildHand))) Like UCase(ChildHand) Then GoTo bone
    Wend
    
    FindChildByClass = 0
    Exit Function
    
bone:
    FindChildByClass = Firs
End Function
Public Sub SetText(Hwnd As Long, Text As String)
    Call SendMessageByString(Hwnd, WM_SETTEXT, 0, Text)
End Sub
Public Function EnumW(Parent As Long, ToDo As Long, All As Boolean) As Boolean
        
    On Local Error GoTo Fehler

    If All Then
        Call EnumWindows(AddressOf WndEnumProc, ToDo)
    Else
        Call EnumChildWindows(ByVal Parent, _
                         AddressOf WndEnumChildProc, ToDo)
    End If
    
    Select Case ToDo
        Case 2
            If Win2Find = -1 Then
                EnumW = True
            Else
                EnumW = False
            End If
            Win2Find = 0
    End Select

Fehler:
End Function
Private Function WndEnumChildProc(ByVal Hwnd As Long, ByVal lParam As Long) As Long
    
    Select Case lParam
        Case 2
            If Hwnd = Win2Find Then
                Win2Find = -1
                WndEnumChildProc = 0
            Else
                WndEnumChildProc = 1
            End If
        Case 3
            Main.Temp.AddItem Hwnd
            WndEnumChildProc = 1
    End Select
    
End Function
Sub FadeForm(Frm As Form, ColStart As Long, ColEnd As Long)
    Dim Red As Single, Green As Single, Blue As Single
    Dim RedStep As Single, GreenStep As Single, BlueStep As Single
    Dim X As Long
    Dim OldSm As Integer
    
    On Local Error GoTo Fehler
    
    Blue = (ColStart \ &H10000) And &HFF
    BlueStep = (Blue - ((ColEnd \ &H10000) And &HFF)) / 64
    Green = (ColStart \ &H100) And &HFF
    GreenStep = (Green - ((ColEnd \ &H100) And &HFF)) / 64
    Red = (ColStart And &HFF)
    RedStep = (Red - (ColEnd And &HFF)) / 64
    
    OldSm = Frm.ScaleMode
    
    Frm.ScaleMode = vbPixels
    Frm.AutoRedraw = True
    Frm.DrawStyle = vbInsideSolid
    Frm.Cls
    Frm.DrawWidth = 2
    Frm.DrawMode = 13
    Frm.ScaleHeight = 64

    For X = 1 To 64
        Frm.Line (0, X)-(Frm.ScaleWidth, X - 1), _
                                         RGB(Red, Green, Blue), BF
        Red = Red - RedStep
        Green = Green - GreenStep
        Blue = Blue - BlueStep
    Next X
    
Fehler:
    Frm.Refresh
    Frm.ScaleMode = OldSm
    
End Sub
Private Function WndEnumProc(ByVal Hwnd As Long, ByVal lParam As Long) As Long
        
    Select Case lParam
        Case 1
            Main.Temp.AddItem Hwnd
            WndEnumProc = 1
        Case 2
            If Hwnd = Win2Find Then
                Win2Find = -1
                WndEnumProc = 0
            Else
                WndEnumProc = 1
            End If
        Case 3
            Main.AllC.AddItem Hwnd
            WndEnumProc = 1
    End Select

End Function
Public Sub AddText(ByRef RTFBox As RichTextBox, _
                   ByVal strText As String, ByVal tColor As Long)
    Dim lngLength As Long, lngSelStart As Long
    
    lngLength = Len(strText)
    lngSelStart = RTFBox.SelStart
    RTFBox.SelLength = 0
    RTFBox.SelText = strText
    RTFBox.SelStart = lngSelStart
    RTFBox.SelLength = Len(strText)
    RTFBox.SelColor = tColor
    RTFBox.SelLength = 0
    RTFBox.SelStart = lngSelStart + Len(strText)
    
End Sub
Public Sub WriteLB(LB As ListBox, RTF As RichTextBox)
    Dim P As Integer
    Dim Color As Long
    Dim LBitem As String
            
    RTF.Text = ""
    
    For P = 0 To LB.ListCount - 1
        LBitem = LB.List(P)
        If IsNumeric(LBitem) Then
            Color = Val(LBitem)
        Else
            AddText RTF, LBitem & vbCrLf, Color
        End If
    Next P

    RTF.SelStart = 0
    RTF.SelLength = 0

End Sub
Public Function GetCinfo(Hwnd As Long) As String
    Dim TaskID As Long
    Dim t As String, V As String
    
    V = "No"
    If IsWindowVisible(Hwnd) Then V = "Yes"
    
    t = GetText(Hwnd)
    If t = "" Then
        t = "Control has no Text"
    Else
        t = Chr(34) & t & Chr(34)
    End If
    
    Call GetWindowThreadProcessId(Hwnd, TaskID)
    
    GetCinfo = "Controlinfo" & vbCrLf & _
               "------------------" & vbCrLf & _
               "   - Hwnd = " & Hwnd & vbCrLf & _
               "   - Classname = " & GetClass(Hwnd) & vbCrLf & _
               "   - TaskID = " & TaskID & vbCrLf & _
               "   - Text = " & t & vbCrLf & _
               "   - Visible = " & V & vbCrLf

End Function
Public Sub PageVisit(URL As String)
    Call ShellExecute(0, "Open", URL, "", "", 1)
End Sub
Public Sub SendMail(Mail As String)
    Dim Buff As String
    
    Buff = "mailto:" & Mail & "?Subject=Re: " & mY.Name & "&Body="
    Call ShellExecute(0&, "Open", Buff, "", "", 1)

End Sub
Public Sub OpenEXE(EXEpath As String)
    Dim Result As Long
    
    Result = ShellExecute(0, "Open", EXEpath, "", "", 1)

    If Result = 2 Then Call uMsg("Path not found:" & vbCrLf & EXEpath, _
                                 "Error", vbCritical + vbOKOnly)
                                 
End Sub
Public Sub LEnum(MyHwnd As Long, Control As Boolean, Childs As Boolean)
    Dim E As New EnumC, A As New Ac, G As New GP
    
    If Control Then
        Load E
        Call E.GetChild(MyHwnd)
    Else
        If Childs Then
            Load A
            Call A.GetHwnd(MyHwnd, False)
        Else
            Load G
            Call G.GetP(MyHwnd)
        End If
    End If
    
End Sub
Public Function FileExist(File2Check As String) As Boolean
    Dim ff As Integer

    On Local Error Resume Next
    
    ff = FreeFile
    
    FileExist = True
    
    Open File2Check For Input As ff
        If Err Then FileExist = False
    Close ff

End Function
Public Sub RollForm(Frm As Form, HeightStart As Integer, HeightEnde As Integer, _
                    WideStart As Integer, WideEnde As Integer, _
                    RollDown As Boolean, Fade As Boolean, _
                    ColStart As Long, ColEnd As Long)
    Dim xStep As Integer
    
    xStep = -150
    If RollDown Then xStep = xStep * (-1)

    For P = WideStart To WideEnde Step xStep
        Frm.Width = P
        If Fade Then
            Call FadeForm(Frm, ColStart, ColEnd)
        Else
            Frm.Refresh
        End If
    Next P
        
    For P = HeightStart To HeightEnde Step xStep
        Frm.Height = P
        If Fade Then
            Call FadeForm(Frm, ColStart, ColEnd)
        Else
            Frm.Refresh
        End If
    Next P

End Sub
Public Function Terminate(Hwnd As Long)
    Dim Result As Long, TaskID As Long
    
    Result = GetWindowThreadProcessId(Hwnd, TaskID)
     
    TaskID = OpenProcess(PROCESS_TERMINATE, 0&, TaskID)
    
    Result = TerminateProcess(TaskID, 1&)
    Result = CloseHandle(TaskID)
    
End Function
Public Sub LBmouseMove(Box As ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim P As Long, LX As Long, LY As Long, Param As Long
    
    LX = Box.Parent.ScaleX(X, Box.Parent.ScaleMode, vbPixels)
    LY = Box.Parent.ScaleY(Y, Box.Parent.ScaleMode, vbPixels)
    Param = CLng(LX) + &H10000 * CLng(LY)

    P = SendMessage(Box.Hwnd, LB_ITEMFROMPOINT, 0, ByVal Param)
    
    If P < Box.ListCount Then
        Box.ListIndex = P
        Box.ToolTipText = Box.List(P)
        If Box.MousePointer <> vbCustom Then
            Box.MouseIcon = fIcon.Hand
            Box.MousePointer = fIcon.Pointer
        End If
    Else
        Box.ListIndex = -1
        If Box.MousePointer <> vbNoDrop Then
            Box.MousePointer = vbNoDrop
            Box.ToolTipText = ""
        End If
    End If
    
End Sub
Public Function getOSversion() As String
    Dim OSVersion As OSVERSIONINFO
    Dim BuildNr As Long
      
    OSVersion.dwOSVersionInfoSize = Len(OSVersion)
    Call GetVersionEx(OSVersion)
    
    With OSVersion
        If (.dwBuildNumber And &HFFFF&) > &H7FFF Then
            BuildNr = (.dwBuildNumber And &HFFFF&) - &H10000
        Else
            BuildNr = .dwBuildNumber And &HFFFF&
        End If
    
        If .dwPlatformId = VER_PLATFORM_WIN32_NT Then
            If .dwMajorVersion = 4 Then
                getOSversion = "Windows NT"
            ElseIf .dwMajorVersion = 5 Then
                If BuildNr = 2600 And .dwMinorVersion > 0 Then
                    getOSversion = "Windows XP"
                Else
                    getOSversion = "Windows 2000"
                End If
            End If
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            If (.dwMajorVersion > 4) Or (.dwMajorVersion = 4 And _
                .dwMinorVersion = 10) Then
                If BuildNr = 1998 Then
                    getOSversion = "Windows 98"
                Else
                    getOSversion = "Windows 98 SE"
                End If
            ElseIf (.dwMajorVersion = 4 And _
                    .dwMinorVersion = 0) Then
                getOSversion = "Windows 95"
            ElseIf (.dwMajorVersion = 4 And _
                    .dwMinorVersion = 90) Then
                getOSversion = "Windows Millenium"
            End If
        ElseIf .dwPlatformId = VER_PLATFORM_WIN32s Then
            getOSversion = "Windows 32s"
        End If
    
    End With
      
End Function
Public Sub GetAllProcess()
    
    Main.LB.Clear
    Main.lbID.Clear
    
    Select Case mY.OS
        Case 1
            Dim F As Long, hSnap As Long
            Dim proc As PROCESSENTRY32
            Dim Sname As String
            
            hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
            
            If hSnap = hNull Then Exit Sub
            
            proc.dwSize = Len(proc)
            
            F = Process32First(hSnap, proc)
            
            Do While F
                Sname = StrZToStr(proc.szExeFile)
                Main.LB.AddItem Sname
                F = Process32Next(hSnap, proc)
            Loop

        Case 2
            Dim CB As Long, cbNeeded As Long
            Dim NumElements As Long, ProcessIDs() As Long
            Dim cbNeeded2 As Long, NumElements2 As Long
            Dim Modules(1 To 200) As Long, lRet As Long
            Dim ModuleName As String
            Dim nSize As Long, hProcess As Long, i As Long
         
            CB = 8
            cbNeeded = 96
         
            Do While CB <= cbNeeded
                CB = CB * 2
                
                ReDim ProcessIDs(CB / 4) As Long
                
                lRet = EnumProcesses(ProcessIDs(1), CB, cbNeeded)
            Loop
         
            NumElements = cbNeeded / 4

            For i = 1 To NumElements
                hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
                    Or PROCESS_VM_READ, 0, ProcessIDs(i))
            
                If hProcess <> 0 Then
                    lRet = EnumProcessModules(hProcess, _
                                              Modules(1), 200, _
                                              cbNeeded2)
                    If lRet <> 0 Then
                        ModuleName = Space(MAX_PATH)
                        nSize = 500
                        lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                                        ModuleName, nSize)
                        Main.LB.AddItem Left(ModuleName, lRet)
                        Main.lbID.AddItem ProcessIDs(i)
                    End If
                End If
                lRet = CloseHandle(hProcess)
            Next i
      End Select
      
End Sub
Private Function StrZToStr(S As String) As String
    StrZToStr = Left$(S, Len(S) - 1)
End Function
Private Function SchuettelnWait(ByVal Delay As Long)
    TickCount = GetTickCount
    
    While (TickCount + Delay) > GetTickCount
        DoEvents
    Wend
    
End Function
Public Function ShakeForm(ByRef fForm As Form, _
                          ByVal lAmplitude As Long, _
                          ByVal lMilliSeconds As Long, _
                          Optional ByVal lFrameRefresh As Long = 10)
 
    Dim lngOriginalLeft As Long, lngOriginalTop As Long
    Dim X As Long, Y As Long
      
    Randomize Timer * Timer
    
    lngOriginalLeft = fForm.Left
    lngOriginalTop = fForm.Top
    
    Do While lMilliSeconds >= lFrameRefresh
        fForm.Left = lngOriginalLeft
        fForm.Top = lngOriginalTop
        
        X = lMilliSeconds / lAmplitude
        Y = lMilliSeconds / lAmplitude
    
        Select Case Int((4) * Rnd + 1)
            Case 1: fForm.Top = fForm.Top - Y
            Case 2: fForm.Top = fForm.Top + Y
            Case 3: fForm.Left = fForm.Left + X
            Case 4: fForm.Left = fForm.Left - X
        End Select
     
        lMilliSeconds = lMilliSeconds - lFrameRefresh
        SchuettelnWait (lFrameRefresh)
    Loop
    
End Function
Public Function LoadStandardForm(Frm As Form, Fade As Boolean)
    Call SetWindowPos(Frm.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
                                           SWP_NOSIZE Or SWP_NOMOVE)
    Frm.Caption = mY.Name
    Frm.MouseIcon = fIcon.Move
    Frm.MousePointer = fIcon.Pointer
    
    Frm.Tag = Frm.Hwnd
    
    If Fade Then Call FadeForm(Frm, 0, 255)
    
End Function
Public Function FindForm(frmName As String, Tag As String) _
                                                        As Long
    Dim P As Integer
    
    frmName = LCase(frmName)
    Tag = LCase(Tag)
    
    Do Until mID(Tag, 1, 1) <> " "
        Tag = mID(Tag, 2)
    Loop
    
    For P = 0 To Forms.Count - 1
        If LCase(Forms(P).Name) = frmName And _
           LCase(Forms(P).Tag) = Tag Then FindForm = P: _
                                          Exit Function
    Next P
    
    FindForm = -1
    
End Function
Public Function SetMP(Obj As Object, Hand As Boolean)
    
    If Hand Then
        Obj.MouseIcon = fIcon.Hand
    Else
        Obj.MouseIcon = fIcon.Move
    End If
    
    Obj.MousePointer = fIcon.Pointer
    
End Function
Public Function Ask(Prompt As String, Title As String, _
                    Button As VbMsgBoxResult) As VbMsgBoxResult
    If Main.PosBorderItem(0).Checked Then Unload DummyDraw
    Main.Hide
    
    Ask = MsgBox(Prompt, Button, Title)

    Main.Show
    
End Function
Public Sub uMsg(Prompt As String, Title As String, _
               Button As VbMsgBoxResult)
    Call Ask(Prompt, Title, Button)
End Sub
Public Function uInput(Prompt As String, Title As String, _
                       Optional Default As String) As String
    
    If Main.PosBorderItem(0).Checked Then Unload DummyDraw
    
    Main.Hide
    
    uInput = InputBox(Prompt, Title, Default)
    
    Main.Show

End Function
Private Function WndProc(ByVal Hwnd As Long, ByVal MSG As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long
    Dim Thief As Long, wMsg As Long, LBindex As Long, Owner As Long
            
    If MSG = WM_COMMAND Then
        wMsg = WM_COMMAND

        Thief = FindForm("Mthief", "mT_" & Hwnd)
                
        LBindex = SearchLB(Forms(Thief).mID.Hwnd, _
                           LB_FINDSTRINGEXACT, -1, wParam)
        Owner = Forms(Thief).Owner.List(LBindex)
        
        If SearchLB(Forms(Thief).SysID.Hwnd, LB_FINDSTRINGEXACT, _
                    -1, wParam) > -1 Then wMsg = WM_SYSCOMMAND
        
        Call PostMessage(Owner, wMsg, wParam, 0&)
    End If
              
    WndProc = CallWindowProc(PrevWndProc, Hwnd, MSG, wParam, lParam)

End Function
Public Function WatchWM(ByVal Hwnd As Long) As Long
    PrevWndProc = SetWindowLong(Hwnd, GWL_WNDPROC, _
                                AddressOf WndProc)
    WatchWM = PrevWndProc
End Function
Public Sub TerminateWatchWM(ByVal Hwnd As Long, PrevWnd As Long)
    Call SetWindowLong(Hwnd, GWL_WNDPROC, PrevWnd)
End Sub
Public Sub WinRepaint(ByVal Hwnd As Long, Show As Boolean)
    Dim R As RECT

    If Show Then Call ShowWindow(Hwnd, SW_SHOW): _
                 Call BringWindowToTop(Hwnd)

    Call GetWindowRect(Hwnd, R)
    
    Call MoveWindow(Hwnd, R.Left, R.Top, _
                          R.Right - R.Left, R.Bottom - R.Top, 1)

End Sub
Public Sub EnumCompleteWindows(LB2Store As ListBox)
    Dim P As Integer, X As Integer
    
    Main.AllC.Clear
    Main.Temp.Clear
    LB2Store.Clear
            
    Call EnumW(0, 3, True)

    For P = 0 To Main.AllC.ListCount - 1
        Main.Temp.Clear
        
        LB2Store.AddItem Main.AllC.List(P)
        Call EnumW(Main.AllC.List(P), 3, False)
    
        For X = 0 To Main.Temp.ListCount - 1
            LB2Store.AddItem Main.Temp.List(X)
        Next X
        
    Next P
    
    Main.AllC.Clear
    
End Sub
Public Sub ChangeWindowStyle(ByVal Hwnd As Long, _
                              ByVal nIndex As Long, Style As Long)
    Dim lngStyle As Long
    Dim R As RECT
      
    lngStyle = GetWindowLong(Hwnd, nIndex)

    If (lngStyle And Style) Then
        lngStyle = lngStyle - Style
    Else
        lngStyle = lngStyle Or Style
    End If
    
    Call GetWindowRect(Hwnd, R)
    
    Call SetWindowLong(Hwnd, nIndex, lngStyle)
    
    Call SetWindowPos(Hwnd, 0, R.Left, R.Top, _
                      R.Right - R.Left, R.Bottom - R.Top, _
                      SWP_FRAMECHANGED)
End Sub
Public Sub FlashWindow(ByVal Hwnd As Long)
    Dim FlashInfo As FLASHWINFO
    
    FlashInfo.cbSize = Len(FlashInfo)
    FlashInfo.dwFlags = FLASHW_CAPTION Or FLASHW_TIMER
    FlashInfo.dwTimeout = 0
    FlashInfo.Hwnd = Hwnd
    FlashInfo.uCount = 2
    
    Call FlashWindowEx(FlashInfo)

End Sub
Public Function GetTopParent(ByVal Hwnd As Long) As Long
    
    GetTopParent = Hwnd
    
    Do Until GetParent(GetTopParent) = 0
        GetTopParent = GetParent(GetTopParent)
    Loop

End Function
Public Function GetWindowState(ByVal Hwnd As Long) As Long
    Dim Style As Long
    
    GetWindowState = SW_NORMAL
    
    Style = GetWindowLongA(Hwnd, GWL_STYLE)
    
    If (Style And WS_MINIMIZE) Then GetWindowState = SW_MINIMIZE
    If (Style And WS_MAXIMIZE) Then GetWindowState = SW_MAXIMIZE
    
End Function
Public Function isUserClick(ByVal Hwnd As Long) As Boolean
    Dim nPoint As POINTAPI
    
    GetCursorPos nPoint
    
    If GetAsyncKeyState(VK_LBUTTON) _
    And WindowFromPoint(nPoint.X, nPoint.Y) = Hwnd Then _
                                                 isUserClick = True

End Function
Public Sub RestoreWindow(ByVal Hwnd As Long, _
                         ByVal SW_STATE As Long, R As RECT)
            
    Call MoveWindow(Hwnd, R.Left, R.Top, _
                          R.Right - R.Left, _
                          R.Bottom - R.Top, 1)

    Call ShowWindow(Hwnd, SW_STATE)
    
End Sub
